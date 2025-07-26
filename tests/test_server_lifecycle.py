"""
Integration tests for server lifecycle management.
"""

import asyncio
import pytest
import tempfile
import os
import signal
import time
from unittest.mock import Mock, patch, AsyncMock
from pathlib import Path

from word_mcp_server.main import ServerManager, main_async, create_default_config
from word_mcp_server.config.config_manager import ConfigManager
from word_mcp_server.utils.setup import SetupManager


class TestServerLifecycle:
    """Test server startup, running, and shutdown lifecycle."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.temp_dir = tempfile.mkdtemp()
        self.config_path = os.path.join(self.temp_dir, "test_config.yaml")
        
        # Create test configuration
        config_manager = ConfigManager()
        config_manager.create_default_config(self.config_path)
    
    def teardown_method(self):
        """Clean up test fixtures."""
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    @pytest.mark.asyncio
    async def test_server_manager_initialization(self):
        """Test ServerManager initialization."""
        server_manager = ServerManager()
        
        assert server_manager.mcp_server is None
        assert server_manager.logger is None
        assert not server_manager.shutdown_event.is_set()
    
    @pytest.mark.asyncio
    async def test_server_manager_signal_handlers(self):
        """Test signal handler setup."""
        server_manager = ServerManager()
        server_manager.logger = Mock()
        
        # Mock signal handling for testing
        with patch('signal.signal') as mock_signal:
            await server_manager.setup_signal_handlers()
            
            # Verify signal handlers were set up
            assert mock_signal.call_count >= 1
    
    @pytest.mark.asyncio
    async def test_server_manager_shutdown(self):
        """Test graceful shutdown process."""
        server_manager = ServerManager()
        server_manager.logger = Mock()
        
        # Mock MCP server with shutdown method
        mock_mcp_server = AsyncMock()
        mock_mcp_server.shutdown = AsyncMock()
        server_manager.mcp_server = mock_mcp_server
        
        # Test shutdown
        await server_manager.shutdown()
        
        # Verify shutdown was called
        mock_mcp_server.shutdown.assert_called_once()
        assert server_manager.shutdown_event.is_set()
        server_manager.logger.info.assert_called()
    
    @pytest.mark.asyncio
    async def test_server_manager_shutdown_with_error(self):
        """Test shutdown handling when MCP server shutdown fails."""
        server_manager = ServerManager()
        server_manager.logger = Mock()
        
        # Mock MCP server that raises exception on shutdown
        mock_mcp_server = AsyncMock()
        mock_mcp_server.shutdown = AsyncMock(side_effect=Exception("Shutdown error"))
        server_manager.mcp_server = mock_mcp_server
        
        # Test shutdown with error
        await server_manager.shutdown()
        
        # Verify error was logged and shutdown event was still set
        server_manager.logger.error.assert_called()
        assert server_manager.shutdown_event.is_set()
    
    @pytest.mark.asyncio
    async def test_server_startup_configuration_loading(self):
        """Test server startup configuration loading phase."""
        server_manager = ServerManager()
        
        with patch('word_mcp_server.main.setup_logging') as mock_setup_logging, \
             patch('word_mcp_server.server.mcp_server.WordMCPServer') as mock_server_class, \
             patch.object(server_manager, 'setup_signal_handlers', new_callable=AsyncMock) as mock_signal_setup:
            
            # Mock MCP server
            mock_server = AsyncMock()
            mock_server.start = AsyncMock()
            mock_server_class.return_value = mock_server
            
            # Mock shutdown event to prevent infinite wait
            async def mock_wait():
                await asyncio.sleep(0.1)  # Short delay
                server_manager.shutdown_event.set()
            
            server_manager.shutdown_event.wait = mock_wait
            
            # Test server startup
            await server_manager.start_server(self.config_path, verbose=True)
            
            # Verify configuration was loaded and logging was set up
            mock_setup_logging.assert_called_once()
            mock_signal_setup.assert_called_once()
            mock_server.start.assert_called_once()
    
    @pytest.mark.asyncio
    async def test_server_startup_with_invalid_config(self):
        """Test server startup with invalid configuration."""
        server_manager = ServerManager()
        
        # Create invalid config file
        invalid_config_path = os.path.join(self.temp_dir, "invalid_config.yaml")
        with open(invalid_config_path, 'w') as f:
            f.write("invalid: yaml: content: [")
        
        # Test startup with invalid config
        with pytest.raises(SystemExit):
            await server_manager.start_server(invalid_config_path)
    
    @pytest.mark.asyncio
    async def test_server_startup_mcp_server_failure(self):
        """Test server startup when MCP server fails to start."""
        server_manager = ServerManager()
        
        with patch('word_mcp_server.main.setup_logging'), \
             patch('word_mcp_server.server.mcp_server.WordMCPServer') as mock_server_class, \
             patch.object(server_manager, 'setup_signal_handlers', new_callable=AsyncMock):
            
            # Mock MCP server that fails to start
            mock_server = AsyncMock()
            mock_server.start = AsyncMock(side_effect=Exception("MCP server start failed"))
            mock_server_class.return_value = mock_server
            
            # Test startup failure
            with pytest.raises(SystemExit):
                await server_manager.start_server(self.config_path)
    
    @pytest.mark.asyncio
    async def test_main_async_function(self):
        """Test main_async function."""
        with patch('word_mcp_server.main.ServerManager') as mock_manager_class:
            mock_manager = AsyncMock()
            mock_manager.start_server = AsyncMock()
            mock_manager_class.return_value = mock_manager
            
            # Test main_async
            await main_async(self.config_path, verbose=True)
            
            # Verify ServerManager was created and started
            mock_manager_class.assert_called_once()
            mock_manager.start_server.assert_called_once_with(self.config_path, True)


class TestConfigurationCreation:
    """Test configuration file creation functionality."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.temp_dir = tempfile.mkdtemp()
        self.config_path = os.path.join(self.temp_dir, "test_config.yaml")
    
    def teardown_method(self):
        """Clean up test fixtures."""
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def test_create_default_config_success(self):
        """Test successful default configuration creation."""
        # Capture stdout
        import io
        from contextlib import redirect_stdout
        
        stdout_capture = io.StringIO()
        
        with redirect_stdout(stdout_capture):
            create_default_config(self.config_path)
        
        # Verify config file was created
        assert os.path.exists(self.config_path)
        
        # Verify output contains expected information
        output = stdout_capture.getvalue()
        assert "Default configuration created" in output
        assert "Configuration summary:" in output
        assert "Server:" in output
        assert "Logging level:" in output
    
    def test_create_default_config_failure(self):
        """Test configuration creation failure."""
        # Try to create config in non-existent directory without permission
        invalid_path = "/root/nonexistent/config.yaml"
        
        with pytest.raises(SystemExit):
            create_default_config(invalid_path)


class TestSetupManager:
    """Test setup and installation utilities."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.setup_manager = SetupManager()
    
    def test_setup_manager_initialization(self):
        """Test SetupManager initialization."""
        assert self.setup_manager.platform in ['windows', 'linux', 'darwin']
        assert isinstance(self.setup_manager.home_dir, Path)
        assert isinstance(self.setup_manager.config_dir, Path)
    
    def test_check_python_packages(self):
        """Test Python package checking."""
        missing_packages = self.setup_manager._check_python_packages()
        
        # Should be a list (may be empty if all packages are installed)
        assert isinstance(missing_packages, list)
        
        # All items should be strings
        for package in missing_packages:
            assert isinstance(package, str)
    
    @patch('platform.system')
    def test_check_system_requirements_non_windows(self, mock_platform):
        """Test system requirements check on non-Windows system."""
        mock_platform.return_value = "Linux"
        
        setup_manager = SetupManager()
        requirements_met, issues = setup_manager.check_system_requirements()
        
        # Should fail on non-Windows due to Word requirement
        assert not requirements_met
        assert any("Windows" in issue for issue in issues)
    
    def test_create_config_directory(self):
        """Test configuration directory creation."""
        import tempfile
        import shutil
        
        # Use temporary directory for testing
        temp_dir = tempfile.mkdtemp()
        
        try:
            setup_manager = SetupManager()
            setup_manager.config_dir = Path(temp_dir) / "test_config"
            
            # Capture stdout
            import io
            from contextlib import redirect_stdout
            
            stdout_capture = io.StringIO()
            
            with redirect_stdout(stdout_capture):
                setup_manager.create_config_directory()
            
            # Verify directories were created
            assert setup_manager.config_dir.exists()
            assert (setup_manager.config_dir / "logs").exists()
            assert (setup_manager.config_dir / "backups").exists()
            assert (setup_manager.config_dir / "temp").exists()
            
            # Verify output
            output = stdout_capture.getvalue()
            assert "Created directory:" in output
            
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)
    
    def test_setup_claude_integration(self):
        """Test Claude MCP configuration generation."""
        import tempfile
        import shutil
        import json
        
        # Use temporary directory for testing
        temp_dir = tempfile.mkdtemp()
        
        try:
            setup_manager = SetupManager()
            setup_manager.config_dir = Path(temp_dir)
            
            config_path = str(Path(temp_dir) / "config.yaml")
            
            # Test Claude integration setup
            claude_config = setup_manager.setup_claude_integration(config_path)
            
            # Verify configuration structure
            assert "mcpServers" in claude_config
            assert "word-office" in claude_config["mcpServers"]
            
            server_config = claude_config["mcpServers"]["word-office"]
            assert server_config["command"] == "python"
            assert "-m" in server_config["args"]
            assert "word_mcp_server" in server_config["args"]
            assert "WORD_MCP_CONFIG" in server_config["env"]
            
            # Verify file was created
            claude_config_path = Path(temp_dir) / "claude_mcp_config.json"
            assert claude_config_path.exists()
            
            # Verify file contents
            with open(claude_config_path, 'r') as f:
                saved_config = json.load(f)
            assert saved_config == claude_config
            
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)
    
    def test_create_startup_scripts(self):
        """Test startup script creation."""
        import tempfile
        import shutil
        
        # Use temporary directory for testing
        temp_dir = tempfile.mkdtemp()
        
        try:
            setup_manager = SetupManager()
            setup_manager.config_dir = Path(temp_dir)
            
            # Test startup script creation
            scripts = setup_manager.create_startup_scripts()
            
            # Verify scripts were created
            assert len(scripts) >= 1  # At least shell script
            
            for script_path in scripts:
                assert os.path.exists(script_path)
                
                # Verify script content
                with open(script_path, 'r') as f:
                    content = f.read()
                assert "word_mcp_server" in content
            
            # Check shell script is executable
            shell_script = Path(temp_dir) / "start_word_mcp_server.sh"
            if shell_script.exists():
                stat = shell_script.stat()
                assert stat.st_mode & 0o111  # Check execute permission
            
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)


class TestMainEntryPoint:
    """Test main entry point command-line handling."""
    
    def test_main_with_setup_command(self):
        """Test main function with --setup command."""
        with patch('sys.argv', ['word-mcp-server', '--setup']), \
             patch('word_mcp_server.utils.setup.SetupManager') as mock_setup_class:
            
            mock_setup_manager = Mock()
            mock_setup_manager.run_full_setup = Mock()
            mock_setup_class.return_value = mock_setup_manager
            
            from word_mcp_server.main import main
            
            # Test main with setup
            main()
            
            # Verify setup was called
            mock_setup_class.assert_called_once()
            mock_setup_manager.run_full_setup.assert_called_once()
    
    def test_main_with_check_requirements_command(self):
        """Test main function with --check-requirements command."""
        with patch('sys.argv', ['word-mcp-server', '--check-requirements']), \
             patch('word_mcp_server.utils.setup.SetupManager') as mock_setup_class:
            
            mock_setup_manager = Mock()
            mock_setup_manager.check_system_requirements = Mock(return_value=(True, []))
            mock_setup_class.return_value = mock_setup_manager
            
            from word_mcp_server.main import main
            
            # Test main with check requirements
            main()
            
            # Verify check was called
            mock_setup_class.assert_called_once()
            mock_setup_manager.check_system_requirements.assert_called_once()
    
    def test_main_with_check_requirements_failure(self):
        """Test main function with --check-requirements when requirements not met."""
        with patch('sys.argv', ['word-mcp-server', '--check-requirements']), \
             patch('word_mcp_server.utils.setup.SetupManager') as mock_setup_class, \
             pytest.raises(SystemExit) as exc_info:
            
            mock_setup_manager = Mock()
            mock_setup_manager.check_system_requirements = Mock(
                return_value=(False, ["Python version too old", "Word not installed"])
            )
            mock_setup_class.return_value = mock_setup_manager
            
            from word_mcp_server.main import main
            
            # Test main with failed requirements check
            main()
        
        # Verify exit code is 1 (failure)
        assert exc_info.value.code == 1
    
    def test_main_with_uninstall_command(self):
        """Test main function with --uninstall command."""
        with patch('sys.argv', ['word-mcp-server', '--uninstall']), \
             patch('word_mcp_server.utils.setup.SetupManager') as mock_setup_class, \
             pytest.raises(SystemExit) as exc_info:
            
            mock_setup_manager = Mock()
            mock_setup_manager.uninstall = Mock(return_value=True)
            mock_setup_class.return_value = mock_setup_manager
            
            from word_mcp_server.main import main
            
            # Test main with uninstall
            main()
        
        # Verify successful exit
        assert exc_info.value.code == 0
        
        # Verify uninstall was called
        mock_setup_class.assert_called_once()
        mock_setup_manager.uninstall.assert_called_once()
    
    def test_main_with_create_config_command(self):
        """Test main function with --create-config command."""
        import tempfile
        import shutil
        
        temp_dir = tempfile.mkdtemp()
        config_path = os.path.join(temp_dir, "test_config.yaml")
        
        try:
            with patch('sys.argv', ['word-mcp-server', '--create-config', '--config-output', config_path]):
                from word_mcp_server.main import main
                
                # Test main with create config
                main()
                
                # Verify config file was created
                assert os.path.exists(config_path)
        
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)
    
    @pytest.mark.asyncio
    async def test_main_with_server_start(self):
        """Test main function starting the server."""
        import tempfile
        import shutil
        
        temp_dir = tempfile.mkdtemp()
        config_path = os.path.join(temp_dir, "test_config.yaml")
        
        # Create test config
        config_manager = ConfigManager()
        config_manager.create_default_config(config_path)
        
        try:
            with patch('sys.argv', ['word-mcp-server', '--config', config_path]), \
                 patch('word_mcp_server.main.main_async', new_callable=AsyncMock) as mock_main_async:
                
                from word_mcp_server.main import main
                
                # Test main with server start
                main()
                
                # Verify main_async was called with correct parameters
                mock_main_async.assert_called_once_with(config_path, False)
        
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)


class TestEnhancedServerLifecycle:
    """Test enhanced server lifecycle features including document cleanup."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.temp_dir = tempfile.mkdtemp()
        self.config_path = os.path.join(self.temp_dir, "test_config.yaml")
        
        # Create test configuration
        config_manager = ConfigManager()
        config_manager.create_default_config(self.config_path)
    
    def teardown_method(self):
        """Clean up test fixtures."""
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    @pytest.mark.asyncio
    async def test_server_shutdown_with_document_cleanup(self):
        """Test server shutdown with document cleanup."""
        server_manager = ServerManager()
        server_manager.logger = Mock()
        
        # Mock MCP server with Word controller
        mock_mcp_server = AsyncMock()
        mock_mcp_server.shutdown = AsyncMock()
        
        # Mock Word controller with cleanup method
        mock_word_controller = AsyncMock()
        mock_word_controller.cleanup_all_documents = AsyncMock(return_value={
            "total_documents": 2,
            "successfully_closed": 2,
            "errors": []
        })
        mock_mcp_server.word_controller = mock_word_controller
        
        server_manager.mcp_server = mock_mcp_server
        
        # Test shutdown with document cleanup
        await server_manager.shutdown()
        
        # Verify document cleanup was called
        mock_word_controller.cleanup_all_documents.assert_called_once()
        mock_mcp_server.shutdown.assert_called_once()
        assert server_manager.shutdown_event.is_set()
        
        # Verify logging messages
        server_manager.logger.info.assert_any_call("Cleaning up documents and Word connections...")
        server_manager.logger.info.assert_any_call("Document cleanup completed")
    
    @pytest.mark.asyncio
    async def test_server_shutdown_with_document_cleanup_failure(self):
        """Test server shutdown when document cleanup fails."""
        server_manager = ServerManager()
        server_manager.logger = Mock()
        
        # Mock MCP server with Word controller
        mock_mcp_server = AsyncMock()
        mock_mcp_server.shutdown = AsyncMock()
        
        # Mock Word controller that fails cleanup
        mock_word_controller = AsyncMock()
        mock_word_controller.cleanup_all_documents = AsyncMock(
            side_effect=Exception("Document cleanup failed")
        )
        mock_mcp_server.word_controller = mock_word_controller
        
        server_manager.mcp_server = mock_mcp_server
        
        # Test shutdown with cleanup failure
        await server_manager.shutdown()
        
        # Verify cleanup was attempted and error was logged
        mock_word_controller.cleanup_all_documents.assert_called_once()
        server_manager.logger.warning.assert_called_with("Document cleanup failed: Document cleanup failed")
        
        # Verify server still shuts down properly
        mock_mcp_server.shutdown.assert_called_once()
        assert server_manager.shutdown_event.is_set()
    
    @pytest.mark.asyncio
    async def test_server_shutdown_without_word_controller(self):
        """Test server shutdown when no Word controller is present."""
        server_manager = ServerManager()
        server_manager.logger = Mock()
        
        # Mock MCP server without Word controller
        mock_mcp_server = AsyncMock()
        mock_mcp_server.shutdown = AsyncMock()
        # No word_controller attribute
        
        server_manager.mcp_server = mock_mcp_server
        
        # Test shutdown without Word controller
        await server_manager.shutdown()
        
        # Verify server shuts down properly without errors
        mock_mcp_server.shutdown.assert_called_once()
        assert server_manager.shutdown_event.is_set()
        server_manager.logger.info.assert_any_call("Graceful shutdown completed")
    
    @pytest.mark.asyncio
    async def test_complete_server_lifecycle_with_cleanup(self):
        """Test complete server lifecycle including startup and shutdown with cleanup."""
        server_manager = ServerManager()
        
        with patch('word_mcp_server.main.setup_logging') as mock_setup_logging, \
             patch('word_mcp_server.server.mcp_server.WordMCPServer') as mock_server_class, \
             patch.object(server_manager, 'setup_signal_handlers', new_callable=AsyncMock) as mock_signal_setup:
            
            # Mock MCP server with Word controller
            mock_server = AsyncMock()
            mock_server.start = AsyncMock()
            mock_server.shutdown = AsyncMock()
            
            mock_word_controller = AsyncMock()
            mock_word_controller.cleanup_all_documents = AsyncMock(return_value={
                "total_documents": 1,
                "successfully_closed": 1,
                "errors": []
            })
            mock_server.word_controller = mock_word_controller
            
            mock_server_class.return_value = mock_server
            
            # Mock shutdown event to simulate server lifecycle
            shutdown_called = False
            
            async def mock_wait():
                nonlocal shutdown_called
                if not shutdown_called:
                    await asyncio.sleep(0.1)  # Simulate server running
                    shutdown_called = True
                    await server_manager.shutdown()  # Trigger shutdown
                await asyncio.sleep(0.1)  # Allow shutdown to complete
            
            server_manager.shutdown_event.wait = mock_wait
            
            # Test complete lifecycle
            await server_manager.start_server(self.config_path, verbose=False)
            
            # Verify startup sequence
            mock_setup_logging.assert_called_once()
            mock_signal_setup.assert_called_once()
            mock_server.start.assert_called_once()
            
            # Verify shutdown sequence
            mock_word_controller.cleanup_all_documents.assert_called_once()
            mock_server.shutdown.assert_called_once()


class TestEnhancedSetupManager:
    """Test enhanced setup manager functionality."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.temp_dir = tempfile.mkdtemp()
        self.setup_manager = SetupManager()
        self.setup_manager.config_dir = Path(self.temp_dir) / "test_config"
    
    def teardown_method(self):
        """Clean up test fixtures."""
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def test_enhanced_setup_with_validation(self):
        """Test enhanced setup process with validation."""
        # Mock system requirements check to pass
        with patch.object(self.setup_manager, 'check_system_requirements', return_value=(True, [])), \
             patch.object(self.setup_manager, '_validate_config_file'), \
             patch.object(self.setup_manager, '_provide_claude_setup_instructions'), \
             patch.object(self.setup_manager, '_validate_startup_scripts'), \
             patch.object(self.setup_manager, '_run_final_validation', return_value={"success": True, "warnings": []}), \
             patch.object(self.setup_manager, '_provide_platform_specific_instructions'):
            
            # Capture stdout
            import io
            from contextlib import redirect_stdout
            
            stdout_capture = io.StringIO()
            
            with redirect_stdout(stdout_capture):
                results = self.setup_manager.run_full_setup()
            
            # Verify setup completed successfully
            assert results["success"] is True
            assert results["setup_phase"] == "completed"
            assert results["config_path"] is not None
            assert results["claude_config"] is not None
            assert len(results["startup_scripts"]) > 0
            
            # Verify output contains expected phases
            output = stdout_capture.getvalue()
            assert "1. Checking system requirements..." in output
            assert "2. Creating configuration directory..." in output
            assert "3. Installing default configuration..." in output
            assert "4. Setting up Claude integration..." in output
            assert "5. Creating startup scripts..." in output
            assert "6. Running final validation..." in output
            assert "Setup completed successfully!" in output
    
    def test_enhanced_setup_with_requirements_failure(self):
        """Test enhanced setup when system requirements fail."""
        # Mock system requirements check to fail
        with patch.object(self.setup_manager, 'check_system_requirements', 
                         return_value=(False, ["Python version too old", "Word not installed"])), \
             patch.object(self.setup_manager, '_provide_requirement_suggestions') as mock_suggestions:
            
            # Capture stdout
            import io
            from contextlib import redirect_stdout
            
            stdout_capture = io.StringIO()
            
            with redirect_stdout(stdout_capture):
                results = self.setup_manager.run_full_setup()
            
            # Verify setup failed at requirements check
            assert results["success"] is False
            assert results["setup_phase"] == "requirements_check"
            assert results["requirements_met"] is False
            assert len(results["issues"]) == 2
            
            # Verify suggestions were provided
            mock_suggestions.assert_called_once()
            
            # Verify output contains failure message
            output = stdout_capture.getvalue()
            assert "System requirements not met:" in output
    
    def test_enhanced_setup_with_permission_error(self):
        """Test enhanced setup when directory creation fails due to permissions."""
        # Mock system requirements check to pass
        with patch.object(self.setup_manager, 'check_system_requirements', return_value=(True, [])), \
             patch.object(self.setup_manager, 'create_config_directory', 
                         side_effect=PermissionError("Permission denied")):
            
            # Capture stdout
            import io
            from contextlib import redirect_stdout
            
            stdout_capture = io.StringIO()
            
            with redirect_stdout(stdout_capture):
                results = self.setup_manager.run_full_setup()
            
            # Verify setup failed at directory creation
            assert results["success"] is False
            assert results["setup_phase"] == "directory_creation"
            
            # Verify helpful error message
            output = stdout_capture.getvalue()
            assert "Permission denied creating configuration directory" in output
            assert "Try running as administrator" in output
    
    def test_enhanced_setup_with_keyboard_interrupt(self):
        """Test enhanced setup when interrupted by user."""
        # Mock system requirements check to pass, then interrupt
        with patch.object(self.setup_manager, 'check_system_requirements', return_value=(True, [])), \
             patch.object(self.setup_manager, 'create_config_directory', 
                         side_effect=KeyboardInterrupt()):
            
            # Capture stdout
            import io
            from contextlib import redirect_stdout
            
            stdout_capture = io.StringIO()
            
            with redirect_stdout(stdout_capture):
                results = self.setup_manager.run_full_setup()
            
            # Verify setup was interrupted
            assert results["success"] is False
            assert results["setup_phase"] == "interrupted"
            
            # Verify interrupt message
            output = stdout_capture.getvalue()
            assert "Setup interrupted by user" in output
    
    def test_requirement_suggestions(self):
        """Test requirement suggestion functionality."""
        issues = [
            "Python 3.7 required, found Python 3.6",
            "Microsoft Word COM automation requires Windows",
            "Microsoft Word not found - please install Microsoft Office",
            "Missing Python packages: mcp, pydantic"
        ]
        
        # Capture stdout
        import io
        from contextlib import redirect_stdout
        
        stdout_capture = io.StringIO()
        
        with redirect_stdout(stdout_capture):
            self.setup_manager._provide_requirement_suggestions(issues)
        
        output = stdout_capture.getvalue()
        
        # Verify suggestions are provided for each issue type
        assert "Update Python" in output
        assert "python.org/downloads" in output
        assert "Windows Subsystem for Linux" in output
        assert "office.microsoft.com" in output
        assert "pip install" in output
    
    def test_config_file_validation(self):
        """Test configuration file validation."""
        # Create a valid config file
        config_path = os.path.join(self.temp_dir, "valid_config.yaml")
        config_manager = ConfigManager()
        config_manager.create_default_config(config_path)
        
        # Test validation of valid config
        try:
            self.setup_manager._validate_config_file(config_path)
            # Should not raise exception
        except Exception as e:
            pytest.fail(f"Valid config validation failed: {e}")
        
        # Test validation of invalid config
        invalid_config_path = os.path.join(self.temp_dir, "invalid_config.yaml")
        with open(invalid_config_path, 'w') as f:
            f.write("invalid: yaml: content: [")
        
        with pytest.raises(Exception) as exc_info:
            self.setup_manager._validate_config_file(invalid_config_path)
        
        assert "Configuration validation failed" in str(exc_info.value)
    
    def test_startup_script_validation(self):
        """Test startup script validation."""
        # Ensure config directory exists
        self.setup_manager.config_dir.mkdir(parents=True, exist_ok=True)
        
        # Create valid startup scripts
        scripts = self.setup_manager.create_startup_scripts()
        
        # Test validation of valid scripts
        try:
            self.setup_manager._validate_startup_scripts(scripts)
            # Should not raise exception
        except Exception as e:
            pytest.fail(f"Valid script validation failed: {e}")
        
        # Test validation of missing script
        missing_script = ["/nonexistent/script.sh"]
        with pytest.raises(Exception) as exc_info:
            self.setup_manager._validate_startup_scripts(missing_script)
        
        assert "Startup script not created" in str(exc_info.value)
    
    def test_final_validation(self):
        """Test final validation process."""
        # Create config file for validation
        config_path = os.path.join(self.temp_dir, "config.yaml")
        config_manager = ConfigManager()
        config_manager.create_default_config(config_path)
        
        # Create required files
        self.setup_manager.config_dir.mkdir(parents=True, exist_ok=True)
        (self.setup_manager.config_dir / "config.yaml").write_text("test")
        (self.setup_manager.config_dir / "claude_mcp_config.json").write_text("{}")
        
        # Test final validation
        results = self.setup_manager._run_final_validation(config_path)
        
        # Should have some validation results
        assert "success" in results
        assert "warnings" in results
        assert isinstance(results["warnings"], list)


class TestWordControllerCleanup:
    """Test Word controller document cleanup functionality."""
    
    def setup_method(self):
        """Set up test fixtures."""
        # Mock Word configuration
        self.mock_config = Mock()
        self.mock_config.save_on_exit = True
    
    @pytest.mark.asyncio
    async def test_cleanup_all_documents_success(self):
        """Test successful document cleanup."""
        # Skip if COM not available (non-Windows)
        try:
            from word_mcp_server.word.controller import WordController, COM_AVAILABLE
            if not COM_AVAILABLE:
                pytest.skip("COM interface not available")
        except ImportError:
            pytest.skip("WordController not available")
        
        with patch('word_mcp_server.word.controller.win32com'), \
             patch('word_mcp_server.word.controller.pythoncom'):
            
            controller = WordController(self.mock_config)
            
            # Mock documents
            mock_doc1 = Mock()
            mock_doc1.Saved = False
            mock_doc1.Save = Mock()
            
            mock_doc2 = Mock()
            mock_doc2.Saved = True
            
            controller._documents = {
                "doc1": Mock(
                    title="Document 1",
                    word_doc_ref=mock_doc1,
                    file_path="/path/to/doc1.docx"
                ),
                "doc2": Mock(
                    title="Document 2", 
                    word_doc_ref=mock_doc2,
                    file_path="/path/to/doc2.docx"
                )
            }
            
            # Mock close_document method
            controller.close_document = Mock()
            
            # Test cleanup
            results = await controller.cleanup_all_documents()
            
            # Verify results
            assert results["total_documents"] == 2
            assert results["successfully_closed"] == 2
            assert len(results["errors"]) == 0
            
            # Verify unsaved document was saved
            mock_doc1.Save.assert_called_once()
            
            # Verify both documents were closed
            assert controller.close_document.call_count == 2
    
    @pytest.mark.asyncio
    async def test_cleanup_all_documents_with_errors(self):
        """Test document cleanup with errors."""
        # Skip if COM not available (non-Windows)
        try:
            from word_mcp_server.word.controller import WordController, COM_AVAILABLE
            if not COM_AVAILABLE:
                pytest.skip("COM interface not available")
        except ImportError:
            pytest.skip("WordController not available")
        
        with patch('word_mcp_server.word.controller.win32com'), \
             patch('word_mcp_server.word.controller.pythoncom'):
            
            controller = WordController(self.mock_config)
            
            # Mock document that fails to save
            mock_doc = Mock()
            mock_doc.Saved = False
            mock_doc.Save = Mock(side_effect=Exception("Save failed"))
            
            controller._documents = {
                "doc1": Mock(
                    title="Document 1",
                    word_doc_ref=mock_doc,
                    file_path="/path/to/doc1.docx"
                )
            }
            
            # Mock close_document method
            controller.close_document = Mock()
            
            # Test cleanup with error
            results = await controller.cleanup_all_documents()
            
            # Verify results include error
            assert results["total_documents"] == 1
            assert results["successfully_closed"] == 1
            assert len(results["errors"]) == 1
            assert "Failed to save document" in results["errors"][0]
            
            # Verify document was still closed despite save error
            controller.close_document.assert_called_once()
    
    @pytest.mark.asyncio
    async def test_cleanup_all_documents_new_document_without_path(self):
        """Test cleanup of new document without file path."""
        # Skip if COM not available (non-Windows)
        try:
            from word_mcp_server.word.controller import WordController, COM_AVAILABLE
            if not COM_AVAILABLE:
                pytest.skip("COM interface not available")
        except ImportError:
            pytest.skip("WordController not available")
        
        with patch('word_mcp_server.word.controller.win32com'), \
             patch('word_mcp_server.word.controller.pythoncom'):
            
            controller = WordController(self.mock_config)
            
            # Mock new document without file path
            mock_doc = Mock()
            mock_doc.Saved = False
            mock_doc.SaveAs2 = Mock()
            
            controller._documents = {
                "doc1": Mock(
                    title="New Document",
                    word_doc_ref=mock_doc,
                    file_path=None  # New document without path
                )
            }
            
            # Mock close_document method
            controller.close_document = Mock()
            
            # Test cleanup
            results = await controller.cleanup_all_documents()
            
            # Verify results
            assert results["total_documents"] == 1
            assert results["successfully_closed"] == 1
            assert len(results["errors"]) == 0
            
            # Verify document was saved with default name
            mock_doc.SaveAs2.assert_called_once()
            call_args = mock_doc.SaveAs2.call_args[0]
            assert "Document_doc1.docx" in call_args[0]