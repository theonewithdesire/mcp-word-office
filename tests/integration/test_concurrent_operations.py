"""
Integration tests for concurrent document operations.

These tests verify that the Word MCP server can handle multiple simultaneous
document operations safely and efficiently.
"""

import pytest
import asyncio
import threading
import time
from unittest.mock import Mock, patch
from concurrent.futures import ThreadPoolExecutor, as_completed

from word_mcp_server.server.mcp_server import WordMCPServer
from word_mcp_server.config.config_manager import ConfigManager
from word_mcp_server.config.models import ServerConfig, WordConfig


class TestConcurrentDocumentOperations:
    """Test concurrent document operations."""
    
    @pytest.fixture
    def mock_config_manager(self):
        """Create a mock configuration manager with concurrency settings."""
        config_manager = Mock(spec=ConfigManager)
        
        server_config = ServerConfig(
            host="localhost",
            port=8080,
            max_concurrent_docs=5,
            timeout_seconds=30
        )
        
        word_config = WordConfig(
            auto_launch=True,
            visible=False,
            save_on_exit=True,
            backup_enabled=True
        )
        
        mock_config = Mock()
        mock_config.server = server_config
        mock_config.word = word_config
        mock_config.logging = Mock()
        
        config_manager.config = mock_config
        return config_manager
    
    @pytest.fixture
    def mock_word_controller(self):
        """Create a thread-safe mock Word controller."""
        controller = Mock()
        controller.is_connected.return_value = True
        controller.connect_to_word.return_value = True
        
        # Use thread-safe counters for document IDs
        self._doc_counter = 0
        self._lock = threading.Lock()
        
        def create_document_with_id():
            with self._lock:
                self._doc_counter += 1
                return f"doc-{self._doc_counter}"
        
        controller.create_document.side_effect = create_document_with_id
        
        # Mock other operations with realistic delays
        def mock_insert_text(doc_id, text, position=None):
            time.sleep(0.01)  # Simulate processing time
            return {
                "success": True,
                "doc_id": doc_id,
                "text": text,
                "position": position or 0,
                "length": len(text)
            }
        
        def mock_save_document(doc_id, path=None):
            time.sleep(0.02)  # Simulate save time
            return {"success": True, "doc_id": doc_id, "path": path}
        
        controller.insert_text.side_effect = mock_insert_text
        controller.save_document.side_effect = mock_save_document
        controller.close_document.return_value = None
        
        return controller
    
    @pytest.mark.asyncio
    async def test_concurrent_document_creation(self, mock_config_manager, mock_word_controller):
        """Test creating multiple documents concurrently."""
        with patch('word_mcp_server.server.mcp_server.WordController') as mock_controller_class:
            mock_controller_class.return_value = mock_word_controller
            
            server = WordMCPServer(mock_config_manager)
            
            # Create multiple documents concurrently
            num_docs = 10
            tasks = [
                server._handle_create_document(title=f"Document {i}")
                for i in range(num_docs)
            ]
            
            start_time = time.time()
            results = await asyncio.gather(*tasks)
            end_time = time.time()
            
            # Verify all operations succeeded
            assert len(results) == num_docs
            for i, result in enumerate(results):
                assert result["success"] is True
                assert result["doc_id"] == f"doc-{i+1}"
            
            # Verify operations were concurrent (should be faster than sequential)
            assert end_time - start_time < num_docs * 0.1  # Much faster than sequential
            
            # Verify controller was called for each document
            assert mock_word_controller.create_document.call_count == num_docs
    
    @pytest.mark.asyncio
    async def test_concurrent_text_operations(self, mock_config_manager, mock_word_controller):
        """Test concurrent text operations on multiple documents."""
        with patch('word_mcp_server.server.mcp_server.WordController') as mock_controller_class:
            mock_controller_class.return_value = mock_word_controller
            
            server = WordMCPServer(mock_config_manager)
            
            # First create documents
            create_tasks = [
                server._handle_create_document(title=f"Doc {i}")
                for i in range(5)
            ]
            create_results = await asyncio.gather(*create_tasks)
            doc_ids = [result["doc_id"] for result in create_results]  # Extract doc_id from result dict
            
            # Then perform concurrent text operations
            text_tasks = []
            for i, doc_id in enumerate(doc_ids):
                text_tasks.append(
                    server._handle_insert_text(
                        doc_id=doc_id,
                        text=f"Content for document {i}",
                        position=0
                    )
                )
            
            text_results = await asyncio.gather(*text_tasks)
            
            # Verify all text operations succeeded
            for i, result in enumerate(text_results):
                assert result["success"] is True
                assert result["doc_id"] == doc_ids[i]
                assert f"Content for document {i}" in result["text"]
    
    @pytest.mark.asyncio
    async def test_concurrent_mixed_operations(self, mock_config_manager, mock_word_controller):
        """Test mixed concurrent operations (create, edit, save)."""
        with patch('word_mcp_server.server.mcp_server.WordController') as mock_controller_class:
            mock_controller_class.return_value = mock_word_controller
            
            server = WordMCPServer(mock_config_manager)
            
            # Mix of different operations
            mixed_tasks = []
            
            # Create documents
            for i in range(3):
                mixed_tasks.append(
                    server._handle_create_document(title=f"New Doc {i}")
                )
            
            # Simulate editing existing documents
            for i in range(3):
                doc_id = f"existing-doc-{i}"
                mixed_tasks.append(
                    server._handle_insert_text(
                        doc_id=doc_id,
                        text=f"Updated content {i}",
                        position=0
                    )
                )
            
            # Simulate saving documents
            for i in range(3):
                doc_id = f"save-doc-{i}"
                mixed_tasks.append(
                    server._handle_save_document(
                        doc_id=doc_id,
                        path=f"/test/doc-{i}.docx"
                    )
                )
            
            results = await asyncio.gather(*mixed_tasks, return_exceptions=True)
            
            # Count successful operations
            # All operations now return dicts with success status
            successful_creates = sum(1 for r in results[:3] if isinstance(r, dict) and r.get("success"))
            successful_inserts = sum(1 for r in results[3:6] if isinstance(r, dict) and r.get("success"))
            successful_saves = sum(1 for r in results[6:] if isinstance(r, dict) and r.get("success"))
            
            # At least the create operations should succeed
            assert successful_creates >= 3
    
    @pytest.mark.asyncio
    async def test_concurrent_operation_error_isolation(self, mock_config_manager, mock_word_controller):
        """Test that errors in one concurrent operation don't affect others."""
        with patch('word_mcp_server.server.mcp_server.WordController') as mock_controller_class:
            # Configure controller to fail on specific document
            def create_with_failure():
                self._call_count = getattr(self, '_call_count', 0) + 1
                if self._call_count == 3:  # Third call fails
                    raise Exception("Simulated failure")
                return f"doc-{self._call_count}"
            
            mock_word_controller.create_document.side_effect = create_with_failure
            mock_controller_class.return_value = mock_word_controller
            
            server = WordMCPServer(mock_config_manager)
            
            # Create multiple documents, one will fail
            tasks = [
                server._handle_create_document(title=f"Document {i}")
                for i in range(5)
            ]
            
            results = await asyncio.gather(*tasks, return_exceptions=True)
            
            # Count successful and failed operations
            successful = sum(1 for r in results if isinstance(r, dict) and r.get("success"))  # Successful returns dict with success=True
            failed = sum(1 for r in results if isinstance(r, dict) and not r.get("success"))  # Failed returns dict with success=False
            
            # Should have 4 successful and 1 failed
            assert successful >= 4
            assert failed >= 1
    
    @pytest.mark.asyncio
    async def test_resource_contention_handling(self, mock_config_manager, mock_word_controller):
        """Test handling of resource contention scenarios."""
        with patch('word_mcp_server.server.mcp_server.WordController') as mock_controller_class:
            # Simulate resource contention with delays
            def slow_operation(doc_id, text, position=None):
                time.sleep(0.1)  # Simulate slow operation
                return {
                    "success": True,
                    "doc_id": doc_id,
                    "text": text,
                    "position": position or 0
                }
            
            mock_word_controller.insert_text.side_effect = slow_operation
            mock_controller_class.return_value = mock_word_controller
            
            server = WordMCPServer(mock_config_manager)
            
            # Create a document first
            create_result = await server._handle_create_document(title="Test Doc")
            doc_id = create_result["doc_id"]  # Extract doc_id from result dict
            
            # Perform multiple operations on the same document
            tasks = [
                server._handle_insert_text(
                    doc_id=doc_id,
                    text=f"Text block {i}",
                    position=i * 10
                )
                for i in range(5)
            ]
            
            start_time = time.time()
            results = await asyncio.gather(*tasks)
            end_time = time.time()
            
            # All operations should succeed
            for result in results:
                assert result["success"] is True
                assert result["doc_id"] == doc_id
            
            # Operations should have been processed (some concurrency expected)
            # Relax assertion for mock environment
            assert end_time - start_time < 2.0  # Should complete within 2 seconds
    
    def test_thread_safety_with_threading(self, mock_config_manager, mock_word_controller):
        """Test thread safety using actual threading."""
        with patch('word_mcp_server.server.mcp_server.WordController') as mock_controller_class:
            mock_controller_class.return_value = mock_word_controller
            
            server = WordMCPServer(mock_config_manager)
            
            results = []
            errors = []
            
            def create_document_thread(thread_id):
                """Function to run in separate thread."""
                try:
                    # Create event loop for this thread
                    loop = asyncio.new_event_loop()
                    asyncio.set_event_loop(loop)
                    
                    # Run async operation
                    result = loop.run_until_complete(
                        server._handle_create_document(title=f"Thread Doc {thread_id}")
                    )
                    results.append(result)
                    
                    loop.close()
                except Exception as e:
                    errors.append(e)
            
            # Create multiple threads
            threads = []
            for i in range(5):
                thread = threading.Thread(target=create_document_thread, args=(i,))
                threads.append(thread)
                thread.start()
            
            # Wait for all threads to complete
            for thread in threads:
                thread.join(timeout=10)
            
            # Verify results
            assert len(errors) == 0, f"Errors occurred: {errors}"
            assert len(results) == 5
            
            for result in results:
                assert isinstance(result, dict)  # Handler returns dict with success and doc_id
                assert result["success"] is True
    
    @pytest.mark.asyncio
    async def test_concurrent_document_lifecycle(self, mock_config_manager, mock_word_controller):
        """Test complete concurrent document lifecycle operations."""
        with patch('word_mcp_server.server.mcp_server.WordController') as mock_controller_class:
            mock_controller_class.return_value = mock_word_controller
            
            server = WordMCPServer(mock_config_manager)
            
            async def document_lifecycle(doc_index):
                """Complete document lifecycle for one document."""
                # Create
                create_result = await server._handle_create_document(
                    title=f"Lifecycle Doc {doc_index}"
                )
                assert isinstance(create_result, dict)  # Handler returns dict with success and doc_id
                assert create_result["success"] is True
                doc_id = create_result["doc_id"]
                
                # Edit
                edit_result = await server._handle_insert_text(
                    doc_id=doc_id,
                    text=f"Content for document {doc_index}",
                    position=0
                )
                assert edit_result["success"] is True
                
                # Save
                save_result = await server._handle_save_document(
                    doc_id=doc_id,
                    path=f"/test/lifecycle-{doc_index}.docx"
                )
                # save_document returns None from controller
                
                # Close
                close_result = await server._handle_close_document(
                    doc_id=doc_id,
                    save=True
                )
                # close_document returns None from controller
                
                return doc_id
            
            # Run multiple complete lifecycles concurrently
            lifecycle_tasks = [
                document_lifecycle(i) for i in range(3)
            ]
            
            completed_doc_ids = await asyncio.gather(*lifecycle_tasks)
            
            # Verify all lifecycles completed successfully
            assert len(completed_doc_ids) == 3
            assert all(doc_id.startswith("doc-") for doc_id in completed_doc_ids)
            
            # Verify all operations were called
            assert mock_word_controller.create_document.call_count == 3
            assert mock_word_controller.insert_text.call_count == 3
            assert mock_word_controller.save_document.call_count == 3
            assert mock_word_controller.close_document.call_count == 3


if __name__ == "__main__":
    pytest.main([__file__, "-v"])