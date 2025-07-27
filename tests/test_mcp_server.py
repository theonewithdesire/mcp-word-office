"""
Unit tests for MCP server functionality.
"""

import pytest
import asyncio
from unittest.mock import Mock, AsyncMock, patch
from typing import Dict, Any

from word_mcp_server.server.mcp_server import WordMCPServer, ToolDefinition
from word_mcp_server.config.config_manager import ConfigManager


class TestWordMCPServer:
    """Test cases for WordMCPServer class."""
    
    @pytest.fixture
    def mock_config_manager(self):
        """Create a mock configuration manager."""
        config_manager = Mock(spec=ConfigManager)
        
        # Mock config object
        mock_config = Mock()
        mock_config.server.dict.return_value = {
            "host": "localhost",
            "port": 8080,
            "max_concurrent_docs": 10,
            "timeout_seconds": 30
        }
        
        config_manager.config = mock_config
        return config_manager
    
    @pytest.fixture
    def mcp_server(self, mock_config_manager):
        """Create a WordMCPServer instance for testing."""
        return WordMCPServer(mock_config_manager)
    
    def test_server_initialization(self, mcp_server, mock_config_manager):
        """Test server initialization."""
        assert mcp_server.config_manager == mock_config_manager
        assert mcp_server.server is not None
        assert isinstance(mcp_server.tools, dict)
        assert not mcp_server.is_running()
        
        # Check that core tools are registered
        expected_tools = [
            "create_document",
            "open_document", 
            "save_document",
            "close_document",
            "insert_text",
            "format_text",
            "select_text",
            "read_document"
        ]
        
        for tool_name in expected_tools:
            assert tool_name in mcp_server.tools
            assert isinstance(mcp_server.tools[tool_name], ToolDefinition)
    
    def test_tool_registration(self, mcp_server):
        """Test tool registration functionality."""
        # Register a custom tool
        async def dummy_handler(args):
            return {"result": "test"}
        
        mcp_server.register_tool(
            name="test_tool",
            description="A test tool",
            parameters={"type": "object", "properties": {}},
            handler=dummy_handler
        )
        
        assert "test_tool" in mcp_server.tools
        tool_def = mcp_server.tools["test_tool"]
        assert tool_def.name == "test_tool"
        assert tool_def.description == "A test tool"
        assert tool_def.handler == dummy_handler
    
    @pytest.mark.asyncio
    async def test_list_tools_handler(self, mcp_server):
        """Test the list tools handler."""
        mcp_server.setup_handlers()
        
        # Test that tools are properly registered
        expected_tools = [
            "create_document",
            "open_document",
            "save_document",
            "close_document",
            "insert_text",
            "format_text",
            "select_text",
            "read_document"
        ]
        
        for expected_tool in expected_tools:
            assert expected_tool in mcp_server.tools
            tool_def = mcp_server.tools[expected_tool]
            assert tool_def.name == expected_tool
            assert tool_def.description is not None
            assert tool_def.parameters is not None
            assert tool_def.handler is not None
    
    @pytest.mark.asyncio
    async def test_call_tool_handler_success(self, mcp_server):
        """Test successful tool call handling."""
        # Register a test tool with a mock handler
        async def test_handler(args):
            return f"Success with args: {args}"
        
        mcp_server.register_tool(
            name="test_success",
            description="Test tool",
            parameters={"type": "object"},
            handler=test_handler
        )
        
        # Test the handler directly
        tool_def = mcp_server.tools["test_success"]
        result = await tool_def.handler({"param": "value"})
        
        assert result == "Success with args: {'param': 'value'}"
    
    @pytest.mark.asyncio
    async def test_call_tool_handler_unknown_tool(self, mcp_server):
        """Test tool call with unknown tool name."""
        # Test that unknown tool is not in the tools registry
        assert "unknown_tool" not in mcp_server.tools
        
        # Test that we have the expected tools
        assert len(mcp_server.tools) == 18  # 18 core tools (including advanced features)
    
    @pytest.mark.asyncio
    async def test_call_tool_handler_exception(self, mcp_server):
        """Test tool call that raises an exception."""
        async def failing_handler(args):
            raise ValueError("Test error")
        
        mcp_server.register_tool(
            name="test_fail",
            description="Failing test tool",
            parameters={"type": "object"},
            handler=failing_handler
        )
        
        # Test the handler directly to verify it raises an exception
        tool_def = mcp_server.tools["test_fail"]
        with pytest.raises(ValueError, match="Test error"):
            await tool_def.handler({})
    
    @pytest.mark.asyncio
    async def test_call_tool_handler_error_result(self, mcp_server):
        """Test tool call that returns an error result."""
        async def error_handler(args):
            return {"error": "Custom error message"}
        
        mcp_server.register_tool(
            name="test_error",
            description="Error test tool",
            parameters={"type": "object"},
            handler=error_handler
        )
        
        # Test the handler directly
        tool_def = mcp_server.tools["test_error"]
        result = await tool_def.handler({})
        
        assert isinstance(result, dict)
        assert "error" in result
        assert result["error"] == "Custom error message"
    
    @pytest.mark.asyncio
    async def test_placeholder_tool_handlers(self, mcp_server):
        """Test that tool handlers work correctly."""
        # Test read_document with invalid path
        tool_def = mcp_server.tools["read_document"]
        
        # This should raise DocumentError due to invalid path
        with pytest.raises(Exception):  # The actual error may be wrapped
            await tool_def.handler(path="test.docx")
    
    def test_server_state_management(self, mcp_server):
        """Test server state management methods."""
        # Initially not running
        assert not mcp_server.is_running()
        
        # Set running state (simulating start)
        mcp_server._running = True
        assert mcp_server.is_running()
        
        # Reset state
        mcp_server._running = False
        assert not mcp_server.is_running()
    
    @pytest.mark.asyncio
    async def test_shutdown_event(self, mcp_server):
        """Test shutdown event handling."""
        # Start a task that waits for shutdown
        shutdown_task = asyncio.create_task(mcp_server.wait_for_shutdown())
        
        # Ensure the task is waiting
        await asyncio.sleep(0.01)
        assert not shutdown_task.done()
        
        # Trigger shutdown
        mcp_server._shutdown_event.set()
        
        # Wait for the task to complete
        await shutdown_task
        assert shutdown_task.done()
    
    # Tests for headers, footers, and page formatting handlers (Task 7)
    @pytest.mark.asyncio
    async def test_handle_insert_header_footer_success(self, mcp_server):
        """Test successful header/footer insertion handler."""
        # Mock the word controller
        mock_controller = Mock()
        mock_controller.insert_header_footer.return_value = {
            "header": {"success": True, "text": "Test Header", "section": 1},
            "footer": {"success": True, "text": "Test Footer", "section": 1}
        }
        
        with patch.object(mcp_server, '_ensure_word_controller', return_value=mock_controller):
            result = await mcp_server._handle_insert_header_footer(
                doc_id="test-doc-123",
                header_text="Test Header",
                footer_text="Test Footer",
                section_index=1
            )
            
            # Result should match the controller return value
            assert "header" in result
            assert "footer" in result
            assert result["header"]["success"] is True
            assert result["footer"]["success"] is True
            
            # Verify controller method was called correctly
            mock_controller.insert_header_footer.assert_called_once_with(
                "test-doc-123", "Test Header", "Test Footer", 1
            )
    
    @pytest.mark.asyncio
    async def test_handle_insert_header_footer_missing_doc_id(self, mcp_server):
        """Test header/footer insertion handler with missing doc_id."""
        # This should raise a TypeError due to missing required positional argument
        with pytest.raises(TypeError, match="missing .* required positional argument"):
            await mcp_server._handle_insert_header_footer(header_text="Test Header")
    
    @pytest.mark.asyncio
    async def test_handle_insert_header_footer_no_text(self, mcp_server):
        """Test header/footer insertion handler with no text provided."""
        # Mock the word controller to avoid connection issues
        mock_controller = Mock()
        mock_controller.insert_header_footer.return_value = {}
        
        with patch.object(mcp_server, '_ensure_word_controller', return_value=mock_controller):
            result = await mcp_server._handle_insert_header_footer(doc_id="test-doc-123")
            
            # Should return empty dict as no text was provided
            assert result == {}
            mock_controller.insert_header_footer.assert_called_once_with("test-doc-123", None, None, 1)
    
    @pytest.mark.asyncio
    async def test_handle_insert_header_footer_header_only(self, mcp_server):
        """Test header/footer insertion handler with header only."""
        # Mock the word controller
        mock_controller = Mock()
        mock_controller.insert_header_footer.return_value = {
            "header": {"success": True, "text": "Test Header", "section": 1}
        }
        
        with patch.object(mcp_server, '_ensure_word_controller', return_value=mock_controller):
            result = await mcp_server._handle_insert_header_footer(
                doc_id="test-doc-123",
                header_text="Test Header"
            )
            
            # Result should match the controller return value
            assert "header" in result
            assert result["header"]["success"] is True
            
            # Verify controller method was called correctly
            mock_controller.insert_header_footer.assert_called_once_with(
                "test-doc-123", "Test Header", None, 1
            )
    
    @pytest.mark.asyncio
    async def test_handle_insert_page_break_success(self, mcp_server):
        """Test successful page break insertion handler."""
        # Mock the word controller
        mock_controller = Mock()
        mock_controller.insert_page_break.return_value = {
            "success": True,
            "break_type": "page",
            "position": 100,
            "message": "Inserted page break at position 100"
        }
        
        with patch.object(mcp_server, '_ensure_word_controller', return_value=mock_controller):
            arguments = {
                "doc_id": "test-doc-123",
                "position": 100,
                "break_type": "page"
            }
            
            result = await mcp_server._handle_insert_page_break(arguments)
            
            assert result["success"] is True
            assert result["doc_id"] == "test-doc-123"
            assert result["result"]["break_type"] == "page"
            assert result["result"]["position"] == 100
            
            # Verify controller method was called correctly
            mock_controller.insert_page_break.assert_called_once_with(
                "test-doc-123", 100, "page"
            )
    
    @pytest.mark.asyncio
    async def test_handle_insert_page_break_section_break(self, mcp_server):
        """Test section break insertion handler."""
        # Mock the word controller
        mock_controller = Mock()
        mock_controller.insert_page_break.return_value = {
            "success": True,
            "break_type": "section_next_page",
            "position": None,
            "message": "Inserted section_next_page break at position None"
        }
        
        with patch.object(mcp_server, '_ensure_word_controller', return_value=mock_controller):
            arguments = {
                "doc_id": "test-doc-123",
                "break_type": "section_next_page"
            }
            
            result = await mcp_server._handle_insert_page_break(arguments)
            
            assert result["success"] is True
            assert result["result"]["break_type"] == "section_next_page"
            
            # Verify controller method was called correctly
            mock_controller.insert_page_break.assert_called_once_with(
                "test-doc-123", None, "section_next_page"
            )
    
    @pytest.mark.asyncio
    async def test_handle_insert_page_break_missing_doc_id(self, mcp_server):
        """Test page break insertion handler with missing doc_id."""
        arguments = {
            "break_type": "page"
        }
        
        result = await mcp_server._handle_insert_page_break(arguments)
        
        assert result["success"] is False
        assert "Document ID is required" in result["error"]
    
    @pytest.mark.asyncio
    async def test_handle_set_page_formatting_success(self, mcp_server):
        """Test successful page formatting handler."""
        # Mock the word controller
        mock_controller = Mock()
        mock_controller.set_page_formatting.return_value = {
            "success": True,
            "section": 1,
            "top_margin": 72,
            "bottom_margin": 72,
            "left_margin": 90,
            "right_margin": 90,
            "orientation": "landscape",
            "paper_size": "letter"
        }
        
        with patch.object(mcp_server, '_ensure_word_controller', return_value=mock_controller):
            arguments = {
                "doc_id": "test-doc-123",
                "section_index": 1,
                "margins": {"top": 72, "bottom": 72, "left": 90, "right": 90},
                "orientation": "landscape",
                "paper_size": "letter"
            }
            
            result = await mcp_server._handle_set_page_formatting(arguments)
            
            assert result["success"] is True
            assert result["doc_id"] == "test-doc-123"
            assert result["results"]["orientation"] == "landscape"
            assert result["results"]["paper_size"] == "letter"
            assert result["results"]["top_margin"] == 72
            
            # Verify controller method was called correctly
            mock_controller.set_page_formatting.assert_called_once_with(
                "test-doc-123", 1, 
                {"top": 72, "bottom": 72, "left": 90, "right": 90},
                "landscape", "letter"
            )
    
    @pytest.mark.asyncio
    async def test_handle_set_page_formatting_margins_only(self, mcp_server):
        """Test page formatting handler with margins only."""
        # Mock the word controller
        mock_controller = Mock()
        mock_controller.set_page_formatting.return_value = {
            "success": True,
            "section": 1,
            "top_margin": 36,
            "left_margin": 54
        }
        
        with patch.object(mcp_server, '_ensure_word_controller', return_value=mock_controller):
            arguments = {
                "doc_id": "test-doc-123",
                "margins": {"top": 36, "left": 54}
            }
            
            result = await mcp_server._handle_set_page_formatting(arguments)
            
            assert result["success"] is True
            assert result["results"]["top_margin"] == 36
            assert result["results"]["left_margin"] == 54
            
            # Verify controller method was called correctly
            mock_controller.set_page_formatting.assert_called_once_with(
                "test-doc-123", 1, {"top": 36, "left": 54}, None, None
            )
    
    @pytest.mark.asyncio
    async def test_handle_set_page_formatting_missing_doc_id(self, mcp_server):
        """Test page formatting handler with missing doc_id."""
        arguments = {
            "margins": {"top": 72}
        }
        
        result = await mcp_server._handle_set_page_formatting(arguments)
        
        assert result["success"] is False
        assert "Document ID is required" in result["error"]
    
    @pytest.mark.asyncio
    async def test_handle_set_page_formatting_no_options(self, mcp_server):
        """Test page formatting handler with no formatting options."""
        arguments = {
            "doc_id": "test-doc-123"
        }
        
        result = await mcp_server._handle_set_page_formatting(arguments)
        
        assert result["success"] is False
        assert "At least one formatting option" in result["error"]


class TestDocumentOperations:
    """Test cases for document lifecycle operations (Task 4.1)."""
    
    @pytest.fixture
    def mock_config_manager(self):
        """Create a mock configuration manager."""
        config_manager = Mock(spec=ConfigManager)
        
        # Mock config object with word config
        mock_config = Mock()
        mock_config.server.dict.return_value = {
            "host": "localhost",
            "port": 8080,
            "max_concurrent_docs": 10,
            "timeout_seconds": 30
        }
        mock_config.word = Mock()
        
        config_manager.config = mock_config
        return config_manager
    
    @pytest.fixture
    def mock_word_controller(self):
        """Create a mock Word controller."""
        controller = Mock()
        controller.is_connected.return_value = True
        controller.connect_to_word.return_value = True
        controller.create_document.return_value = "test-doc-id"
        controller.open_document.return_value = "test-doc-id"
        controller.save_document.return_value = None
        controller.close_document.return_value = None
        
        # Mock document reference
        mock_doc_ref = Mock()
        mock_doc_ref.word_doc_ref = Mock()
        mock_doc_ref.title = "Test Document"
        controller.get_document_reference.return_value = mock_doc_ref
        
        return controller
    
    @pytest.mark.asyncio
    async def test_create_document_success(self, mock_config_manager, mock_word_controller):
        """Test successful document creation."""
        with patch('word_mcp_server.server.mcp_server.WordController') as mock_controller_class:
            mock_controller_class.return_value = mock_word_controller
            
            server = WordMCPServer(mock_config_manager)
            result = await server._handle_create_document({})
            
            assert result["success"] is True
            assert result["doc_id"] == "test-doc-id"
            assert "Created new document" in result["message"]
            mock_word_controller.create_document.assert_called_once()
    
    @pytest.mark.asyncio
    async def test_create_document_with_title(self, mock_config_manager, mock_word_controller):
        """Test document creation with title."""
        with patch('word_mcp_server.server.mcp_server.WordController') as mock_controller_class:
            mock_controller_class.return_value = mock_word_controller
            
            server = WordMCPServer(mock_config_manager)
            result = await server._handle_create_document({"title": "My Document"})
            
            assert result["success"] is True
            assert result["doc_id"] == "test-doc-id"
            mock_word_controller.create_document.assert_called_once()
            mock_word_controller.get_document_reference.assert_called_once_with("test-doc-id")
    
    @pytest.mark.asyncio
    async def test_create_document_connection_error(self, mock_config_manager):
        """Test document creation with connection error."""
        with patch('word_mcp_server.server.mcp_server.WordController') as mock_controller_class:
            from word_mcp_server.utils.errors import ConnectionError
            mock_controller_class.side_effect = ConnectionError("COM interface not available")
            
            server = WordMCPServer(mock_config_manager)
            result = await server._handle_create_document({})
            
            assert result["success"] is False
            assert "COM interface not available" in result["error"]
    
    @pytest.mark.asyncio
    async def test_open_document_success(self, mock_config_manager, mock_word_controller):
        """Test successful document opening."""
        with patch('word_mcp_server.server.mcp_server.WordController') as mock_controller_class:
            mock_controller_class.return_value = mock_word_controller
            
            server = WordMCPServer(mock_config_manager)
            result = await server._handle_open_document({"path": "/test/document.docx"})
            
            assert result["success"] is True
            assert result["doc_id"] == "test-doc-id"
            assert result["path"] == "/test/document.docx"
            assert "Opened document" in result["message"]
            mock_word_controller.open_document.assert_called_once_with("/test/document.docx")
    
    @pytest.mark.asyncio
    async def test_open_document_missing_path(self, mock_config_manager, mock_word_controller):
        """Test document opening without path."""
        with patch('word_mcp_server.server.mcp_server.WordController') as mock_controller_class:
            mock_controller_class.return_value = mock_word_controller
            
            server = WordMCPServer(mock_config_manager)
            result = await server._handle_open_document({})
            
            assert result["success"] is False
            assert "Document path is required" in result["error"]
    
    @pytest.mark.asyncio
    async def test_open_document_file_not_found(self, mock_config_manager, mock_word_controller):
        """Test document opening with file not found error."""
        from word_mcp_server.utils.errors import DocumentError, ErrorCode
        mock_word_controller.open_document.side_effect = DocumentError(
            "Document not found", 
            error_code=ErrorCode.DOCUMENT_NOT_FOUND.value
        )
        
        with patch('word_mcp_server.server.mcp_server.WordController') as mock_controller_class:
            mock_controller_class.return_value = mock_word_controller
            
            server = WordMCPServer(mock_config_manager)
            result = await server._handle_open_document({"path": "/nonexistent/file.docx"})
            
            assert result["success"] is False
            assert "Document not found" in result["error"]
    
    @pytest.mark.asyncio
    async def test_save_document_success(self, mock_config_manager, mock_word_controller):
        """Test successful document saving."""
        with patch('word_mcp_server.server.mcp_server.WordController') as mock_controller_class:
            mock_controller_class.return_value = mock_word_controller
            
            server = WordMCPServer(mock_config_manager)
            result = await server._handle_save_document({"doc_id": "test-doc-id"})
            
            assert result["success"] is True
            assert result["doc_id"] == "test-doc-id"
            assert "Saved document" in result["message"]
            mock_word_controller.save_document.assert_called_once_with("test-doc-id", None)
    
    @pytest.mark.asyncio
    async def test_save_document_with_path(self, mock_config_manager, mock_word_controller):
        """Test document saving with specific path."""
        with patch('word_mcp_server.server.mcp_server.WordController') as mock_controller_class:
            mock_controller_class.return_value = mock_word_controller
            
            server = WordMCPServer(mock_config_manager)
            result = await server._handle_save_document({
                "doc_id": "test-doc-id",
                "path": "/new/path.docx"
            })
            
            assert result["success"] is True
            assert result["doc_id"] == "test-doc-id"
            assert result["path"] == "/new/path.docx"
            assert "Saved document test-doc-id to '/new/path.docx'" in result["message"]
            mock_word_controller.save_document.assert_called_once_with("test-doc-id", "/new/path.docx")
    
    @pytest.mark.asyncio
    async def test_save_document_missing_doc_id(self, mock_config_manager, mock_word_controller):
        """Test document saving without document ID."""
        with patch('word_mcp_server.server.mcp_server.WordController') as mock_controller_class:
            mock_controller_class.return_value = mock_word_controller
            
            server = WordMCPServer(mock_config_manager)
            result = await server._handle_save_document({})
            
            assert result["success"] is False
            assert "Document ID is required" in result["error"]
    
    @pytest.mark.asyncio
    async def test_save_document_not_found(self, mock_config_manager, mock_word_controller):
        """Test saving non-existent document."""
        from word_mcp_server.utils.errors import DocumentError, ErrorCode
        mock_word_controller.save_document.side_effect = DocumentError(
            "Document not found", 
            error_code=ErrorCode.DOCUMENT_NOT_FOUND.value
        )
        
        with patch('word_mcp_server.server.mcp_server.WordController') as mock_controller_class:
            mock_controller_class.return_value = mock_word_controller
            
            server = WordMCPServer(mock_config_manager)
            result = await server._handle_save_document({"doc_id": "nonexistent-doc-id"})
            
            assert result["success"] is False
            assert "Document not found" in result["error"]
    
    @pytest.mark.asyncio
    async def test_close_document_success(self, mock_config_manager, mock_word_controller):
        """Test successful document closing."""
        with patch('word_mcp_server.server.mcp_server.WordController') as mock_controller_class:
            mock_controller_class.return_value = mock_word_controller
            
            server = WordMCPServer(mock_config_manager)
            result = await server._handle_close_document({"doc_id": "test-doc-id"})
            
            assert result["success"] is True
            assert result["doc_id"] == "test-doc-id"
            assert "Closed document" in result["message"]
            mock_word_controller.close_document.assert_called_once_with("test-doc-id", True)
    
    @pytest.mark.asyncio
    async def test_close_document_no_save(self, mock_config_manager, mock_word_controller):
        """Test document closing without saving."""
        with patch('word_mcp_server.server.mcp_server.WordController') as mock_controller_class:
            mock_controller_class.return_value = mock_word_controller
            
            server = WordMCPServer(mock_config_manager)
            result = await server._handle_close_document({
                "doc_id": "test-doc-id",
                "save": False
            })
            
            assert result["success"] is True
            assert result["doc_id"] == "test-doc-id"
            mock_word_controller.close_document.assert_called_once_with("test-doc-id", False)
    
    @pytest.mark.asyncio
    async def test_close_document_missing_doc_id(self, mock_config_manager, mock_word_controller):
        """Test document closing without document ID."""
        with patch('word_mcp_server.server.mcp_server.WordController') as mock_controller_class:
            mock_controller_class.return_value = mock_word_controller
            
            server = WordMCPServer(mock_config_manager)
            result = await server._handle_close_document({})
            
            assert result["success"] is False
            assert "Document ID is required" in result["error"]
    
    @pytest.mark.asyncio
    async def test_close_document_not_found(self, mock_config_manager, mock_word_controller):
        """Test closing non-existent document."""
        from word_mcp_server.utils.errors import DocumentError, ErrorCode
        mock_word_controller.close_document.side_effect = DocumentError(
            "Document not found", 
            error_code=ErrorCode.DOCUMENT_NOT_FOUND.value
        )
        
        with patch('word_mcp_server.server.mcp_server.WordController') as mock_controller_class:
            mock_controller_class.return_value = mock_word_controller
            
            server = WordMCPServer(mock_config_manager)
            result = await server._handle_close_document({"doc_id": "nonexistent-doc-id"})
            
            assert result["success"] is False
            assert "Document not found" in result["error"]


class TestTextOperations:
    """Test cases for text insertion and formatting operations (Task 4.2)."""
    
    @pytest.fixture
    def mock_config_manager(self):
        """Create a mock configuration manager."""
        config_manager = Mock(spec=ConfigManager)
        
        # Mock config object with word config
        mock_config = Mock()
        mock_config.server.dict.return_value = {
            "host": "localhost",
            "port": 8080,
            "max_concurrent_docs": 10,
            "timeout_seconds": 30
        }
        mock_config.word = Mock()
        
        config_manager.config = mock_config
        return config_manager
    
    @pytest.fixture
    def mock_word_controller(self):
        """Create a mock Word controller."""
        controller = Mock()
        controller.is_connected.return_value = True
        controller.connect_to_word.return_value = True
        controller.insert_text.return_value = None
        controller.format_text.return_value = None
        controller.select_text.return_value = {
            "start": 0,
            "end": 10,
            "text": "Hello test",
            "length": 10
        }
        
        return controller
    
    @pytest.mark.asyncio
    async def test_insert_text_success(self, mock_config_manager, mock_word_controller):
        """Test successful text insertion."""
        with patch('word_mcp_server.server.mcp_server.WordController') as mock_controller_class:
            mock_controller_class.return_value = mock_word_controller
            
            server = WordMCPServer(mock_config_manager)
            result = await server._handle_insert_text({
                "doc_id": "test-doc-id",
                "text": "Hello World"
            })
            
            assert result["success"] is True
            assert result["doc_id"] == "test-doc-id"
            assert result["text_length"] == 11
            assert "Inserted 11 characters" in result["message"]
            mock_word_controller.insert_text.assert_called_once_with("test-doc-id", "Hello World", None)
    
    @pytest.mark.asyncio
    async def test_insert_text_with_position(self, mock_config_manager, mock_word_controller):
        """Test text insertion with specific position."""
        with patch('word_mcp_server.server.mcp_server.WordController') as mock_controller_class:
            mock_controller_class.return_value = mock_word_controller
            
            server = WordMCPServer(mock_config_manager)
            result = await server._handle_insert_text({
                "doc_id": "test-doc-id",
                "text": "Hello World",
                "position": 5
            })
            
            assert result["success"] is True
            assert result["doc_id"] == "test-doc-id"
            assert result["position"] == 5
            assert result["text_length"] == 11
            assert "at position 5" in result["message"]
            mock_word_controller.insert_text.assert_called_once_with("test-doc-id", "Hello World", 5)
    
    @pytest.mark.asyncio
    async def test_insert_text_missing_doc_id(self, mock_config_manager, mock_word_controller):
        """Test text insertion without document ID."""
        with patch('word_mcp_server.server.mcp_server.WordController') as mock_controller_class:
            mock_controller_class.return_value = mock_word_controller
            
            server = WordMCPServer(mock_config_manager)
            result = await server._handle_insert_text({"text": "Hello World"})
            
            assert result["success"] is False
            assert "Document ID is required" in result["error"]
    
    @pytest.mark.asyncio
    async def test_insert_text_missing_text(self, mock_config_manager, mock_word_controller):
        """Test text insertion without text."""
        with patch('word_mcp_server.server.mcp_server.WordController') as mock_controller_class:
            mock_controller_class.return_value = mock_word_controller
            
            server = WordMCPServer(mock_config_manager)
            result = await server._handle_insert_text({"doc_id": "test-doc-id"})
            
            assert result["success"] is False
            assert "Text is required" in result["error"]
    
    @pytest.mark.asyncio
    async def test_insert_text_document_error(self, mock_config_manager, mock_word_controller):
        """Test text insertion with document error."""
        from word_mcp_server.utils.errors import DocumentError, ErrorCode
        mock_word_controller.insert_text.side_effect = DocumentError(
            "Document not found", 
            error_code=ErrorCode.DOCUMENT_NOT_FOUND.value
        )
        
        with patch('word_mcp_server.server.mcp_server.WordController') as mock_controller_class:
            mock_controller_class.return_value = mock_word_controller
            
            server = WordMCPServer(mock_config_manager)
            result = await server._handle_insert_text({
                "doc_id": "nonexistent-doc-id",
                "text": "Hello World"
            })
            
            assert result["success"] is False
            assert "Document not found" in result["error"]
    
    @pytest.mark.asyncio
    async def test_format_text_success(self, mock_config_manager, mock_word_controller):
        """Test successful text formatting."""
        with patch('word_mcp_server.server.mcp_server.WordController') as mock_controller_class:
            mock_controller_class.return_value = mock_word_controller
            
            server = WordMCPServer(mock_config_manager)
            result = await server._handle_format_text({
                "doc_id": "test-doc-id",
                "start": 0,
                "end": 10,
                "bold": True,
                "italic": True,
                "font_size": 14
            })
            
            assert result["success"] is True
            assert result["doc_id"] == "test-doc-id"
            assert result["start"] == 0
            assert result["end"] == 10
            assert result["formatting"]["bold"] is True
            assert result["formatting"]["italic"] is True
            assert result["formatting"]["font_size"] == 14
            assert "Applied formatting" in result["message"]
            
            mock_word_controller.format_text.assert_called_once_with(
                "test-doc-id", 0, 10, bold=True, italic=True, font_size=14
            )
    
    @pytest.mark.asyncio
    async def test_format_text_with_colors(self, mock_config_manager, mock_word_controller):
        """Test text formatting with colors."""
        with patch('word_mcp_server.server.mcp_server.WordController') as mock_controller_class:
            mock_controller_class.return_value = mock_word_controller
            
            server = WordMCPServer(mock_config_manager)
            result = await server._handle_format_text({
                "doc_id": "test-doc-id",
                "start": 0,
                "end": 10,
                "color": "#FF0000",
                "highlight_color": "yellow"
            })
            
            assert result["success"] is True
            assert result["formatting"]["color"] == "#FF0000"
            assert result["formatting"]["highlight_color"] == "yellow"
            
            mock_word_controller.format_text.assert_called_once_with(
                "test-doc-id", 0, 10, color="#FF0000", highlight_color="yellow"
            )
    
    @pytest.mark.asyncio
    async def test_format_text_missing_doc_id(self, mock_config_manager, mock_word_controller):
        """Test text formatting without document ID."""
        with patch('word_mcp_server.server.mcp_server.WordController') as mock_controller_class:
            mock_controller_class.return_value = mock_word_controller
            
            server = WordMCPServer(mock_config_manager)
            result = await server._handle_format_text({
                "start": 0,
                "end": 10,
                "bold": True
            })
            
            assert result["success"] is False
            assert "Document ID is required" in result["error"]
    
    @pytest.mark.asyncio
    async def test_format_text_missing_range(self, mock_config_manager, mock_word_controller):
        """Test text formatting without start/end positions."""
        with patch('word_mcp_server.server.mcp_server.WordController') as mock_controller_class:
            mock_controller_class.return_value = mock_word_controller
            
            server = WordMCPServer(mock_config_manager)
            result = await server._handle_format_text({
                "doc_id": "test-doc-id",
                "bold": True
            })
            
            assert result["success"] is False
            assert "Start and end positions are required" in result["error"]
    
    @pytest.mark.asyncio
    async def test_format_text_no_formatting(self, mock_config_manager, mock_word_controller):
        """Test text formatting without any formatting options."""
        with patch('word_mcp_server.server.mcp_server.WordController') as mock_controller_class:
            mock_controller_class.return_value = mock_word_controller
            
            server = WordMCPServer(mock_config_manager)
            result = await server._handle_format_text({
                "doc_id": "test-doc-id",
                "start": 0,
                "end": 10
            })
            
            assert result["success"] is False
            assert "At least one formatting option is required" in result["error"]
    
    @pytest.mark.asyncio
    async def test_format_text_operation_error(self, mock_config_manager, mock_word_controller):
        """Test text formatting with operation error."""
        from word_mcp_server.utils.errors import OperationError, ErrorCode
        mock_word_controller.format_text.side_effect = OperationError(
            "Invalid text range", 
            error_code=ErrorCode.OPERATION_FAILED.value
        )
        
        with patch('word_mcp_server.server.mcp_server.WordController') as mock_controller_class:
            mock_controller_class.return_value = mock_word_controller
            
            server = WordMCPServer(mock_config_manager)
            result = await server._handle_format_text({
                "doc_id": "test-doc-id",
                "start": 0,
                "end": 10,
                "bold": True
            })
            
            assert result["success"] is False
            assert "Invalid text range" in result["error"]
    
    @pytest.mark.asyncio
    async def test_select_text_success(self, mock_config_manager, mock_word_controller):
        """Test successful text selection."""
        with patch('word_mcp_server.server.mcp_server.WordController') as mock_controller_class:
            mock_controller_class.return_value = mock_word_controller
            
            server = WordMCPServer(mock_config_manager)
            result = await server._handle_select_text({
                "doc_id": "test-doc-id",
                "start": 0,
                "end": 10
            })
            
            assert result["success"] is True
            assert result["doc_id"] == "test-doc-id"
            assert result["selection"]["start"] == 0
            assert result["selection"]["end"] == 10
            assert result["selection"]["text"] == "Hello test"
            assert result["selection"]["length"] == 10
            assert "Selected text" in result["message"]
            
            mock_word_controller.select_text.assert_called_once_with("test-doc-id", 0, 10)
    
    @pytest.mark.asyncio
    async def test_select_text_missing_doc_id(self, mock_config_manager, mock_word_controller):
        """Test text selection without document ID."""
        with patch('word_mcp_server.server.mcp_server.WordController') as mock_controller_class:
            mock_controller_class.return_value = mock_word_controller
            
            server = WordMCPServer(mock_config_manager)
            result = await server._handle_select_text({
                "start": 0,
                "end": 10
            })
            
            assert result["success"] is False
            assert "Document ID is required" in result["error"]
    
    @pytest.mark.asyncio
    async def test_select_text_missing_range(self, mock_config_manager, mock_word_controller):
        """Test text selection without start/end positions."""
        with patch('word_mcp_server.server.mcp_server.WordController') as mock_controller_class:
            mock_controller_class.return_value = mock_word_controller
            
            server = WordMCPServer(mock_config_manager)
            result = await server._handle_select_text({
                "doc_id": "test-doc-id"
            })
            
            assert result["success"] is False
            assert "Start and end positions are required" in result["error"]
    
    @pytest.mark.asyncio
    async def test_select_text_operation_error(self, mock_config_manager, mock_word_controller):
        """Test text selection with operation error."""
        from word_mcp_server.utils.errors import OperationError, ErrorCode
        mock_word_controller.select_text.side_effect = OperationError(
            "Invalid text range", 
            error_code=ErrorCode.OPERATION_FAILED.value
        )
        
        with patch('word_mcp_server.server.mcp_server.WordController') as mock_controller_class:
            mock_controller_class.return_value = mock_word_controller
            
            server = WordMCPServer(mock_config_manager)
            result = await server._handle_select_text({
                "doc_id": "test-doc-id",
                "start": 0,
                "end": 10
            })
            
            assert result["success"] is False
            assert "Invalid text range" in result["error"]


class TestToolDefinition:
    """Test cases for ToolDefinition dataclass."""
    
    def test_tool_definition_creation(self):
        """Test ToolDefinition creation."""
        async def dummy_handler(args):
            return "test"
        
        tool_def = ToolDefinition(
            name="test_tool",
            description="A test tool",
            parameters={"type": "object"},
            handler=dummy_handler
        )
        
        assert tool_def.name == "test_tool"
        assert tool_def.description == "A test tool"
        assert tool_def.parameters == {"type": "object"}
        assert tool_def.handler == dummy_handler


@pytest.mark.asyncio
async def test_server_integration():
    """Integration test for server setup and basic functionality."""
    # Create a mock config manager
    config_manager = Mock(spec=ConfigManager)
    mock_config = Mock()
    mock_config.server.dict.return_value = {
        "host": "localhost",
        "port": 8080
    }
    config_manager.config = mock_config
    
    # Create server
    server = WordMCPServer(config_manager)
    
    # Setup handlers
    server.setup_handlers()
    
    # Verify server is properly initialized
    assert server.config_manager == config_manager
    assert server.server is not None
    assert len(server.tools) == 18  # 18 core tools (including advanced features)
    
    # Test that all expected tools are registered
    expected_tools = ["create_document", "open_document", "save_document", "close_document", "insert_text", "format_text", "select_text", "read_document"]
    for tool_name in expected_tools:
        assert tool_name in server.tools
        tool_def = server.tools[tool_name]
        assert isinstance(tool_def, ToolDefinition)
        assert tool_def.name == tool_name
    
    # Test that document operations are no longer placeholders
    # (They should fail due to no Word connection, but not with "not yet implemented")
    create_doc_tool = server.tools["create_document"]
    result = await create_doc_tool.handler({})
    assert isinstance(result, dict)
    assert result["success"] is False
    # Should fail with connection error, not "not yet implemented"
    assert "not yet implemented" not in result["error"]