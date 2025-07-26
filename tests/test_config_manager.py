"""
Tests for configuration manager.
"""

import os
import tempfile
import pytest
import yaml
from pathlib import Path
from pydantic import ValidationError

from word_mcp_server.config.config_manager import ConfigManager
from word_mcp_server.config.models import AppConfig, ServerConfig, WordConfig, LoggingConfig, SecurityConfig


class TestConfigManager:
    """Test configuration manager functionality."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.temp_dir = tempfile.mkdtemp()
        self.config_path = os.path.join(self.temp_dir, "test_config.yaml")
        
    def teardown_method(self):
        """Clean up test fixtures."""
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def test_load_default_config(self):
        """Test loading default configuration when no file exists."""
        config_manager = ConfigManager()
        config = config_manager.load_config()
        
        assert isinstance(config, AppConfig)
        assert config.server.host == "localhost"
        assert config.server.port == 8080
        assert config.word.auto_launch is True
        assert config.logging.level == "INFO"
        assert config.security.enable_macros is False
    
    def test_load_config_from_file(self):
        """Test loading configuration from YAML file."""
        config_data = {
            "server": {
                "host": "0.0.0.0",
                "port": 9090,
                "max_concurrent_docs": 5
            },
            "word": {
                "auto_launch": False,
                "visible": True
            }
        }
        
        with open(self.config_path, 'w') as f:
            yaml.dump(config_data, f)
        
        config_manager = ConfigManager(self.config_path)
        config = config_manager.load_config()
        
        assert config.server.host == "0.0.0.0"
        assert config.server.port == 9090
        assert config.server.max_concurrent_docs == 5
        assert config.word.auto_launch is False
        assert config.word.visible is True
    
    def test_load_config_file_not_found(self):
        """Test error when specified config file doesn't exist."""
        config_manager = ConfigManager("/nonexistent/config.yaml")
        
        with pytest.raises(FileNotFoundError, match="Configuration file not found"):
            config_manager.load_config()
    
    def test_load_config_invalid_yaml(self):
        """Test error when YAML file is malformed."""
        with open(self.config_path, 'w') as f:
            f.write("invalid: yaml: content: [")
        
        config_manager = ConfigManager(self.config_path)
        
        with pytest.raises(yaml.YAMLError, match="Failed to parse YAML file"):
            config_manager.load_config()
    
    def test_load_config_validation_error(self):
        """Test helpful error messages for validation failures."""
        config_data = {
            "server": {
                "port": 70000,  # Invalid port
                "max_concurrent_docs": -1  # Invalid value
            },
            "logging": {
                "level": "INVALID"  # Invalid log level
            }
        }
        
        with open(self.config_path, 'w') as f:
            yaml.dump(config_data, f)
        
        config_manager = ConfigManager(self.config_path)
        
        with pytest.raises(ValueError) as exc_info:
            config_manager.load_config()
        
        error_message = str(exc_info.value)
        assert "Configuration validation failed" in error_message
        assert "server.port" in error_message
        assert "server.max_concurrent_docs" in error_message
        assert "logging.level" in error_message
        assert "python -m word_mcp_server --create-config" in error_message
    
    def test_save_config(self):
        """Test saving configuration to file."""
        config_manager = ConfigManager()
        config = config_manager.load_config()
        
        # Modify some values
        config.server.port = 9999
        config.word.visible = True
        
        config_manager.save_config(self.config_path)
        
        # Verify file was created and contains correct data
        assert os.path.exists(self.config_path)
        
        with open(self.config_path, 'r') as f:
            saved_data = yaml.safe_load(f)
        
        assert saved_data['server']['port'] == 9999
        assert saved_data['word']['visible'] is True
    
    def test_save_config_no_config_loaded(self):
        """Test error when trying to save without loaded config."""
        config_manager = ConfigManager()
        
        with pytest.raises(ValueError, match="No configuration loaded to save"):
            config_manager.save_config(self.config_path)
    
    def test_create_default_config(self):
        """Test creating default configuration file."""
        config_manager = ConfigManager()
        config = config_manager.create_default_config(self.config_path)
        
        assert isinstance(config, AppConfig)
        assert os.path.exists(self.config_path)
        
        # Verify file contains default values
        with open(self.config_path, 'r') as f:
            saved_data = yaml.safe_load(f)
        
        assert saved_data['server']['host'] == "localhost"
        assert saved_data['server']['port'] == 8080
        assert saved_data['word']['auto_launch'] is True
        assert saved_data['logging']['level'] == "INFO"
    
    def test_create_default_config_creates_directory(self):
        """Test that creating default config creates necessary directories."""
        nested_path = os.path.join(self.temp_dir, "nested", "dir", "config.yaml")
        config_manager = ConfigManager()
        
        config_manager.create_default_config(nested_path)
        
        assert os.path.exists(nested_path)
        assert os.path.isdir(os.path.dirname(nested_path))
    
    def test_config_property_lazy_loading(self):
        """Test that config property loads configuration lazily."""
        config_manager = ConfigManager()
        
        # Config should not be loaded initially
        assert config_manager._config is None
        
        # Accessing config property should load it
        config = config_manager.config
        assert isinstance(config, AppConfig)
        assert config_manager._config is not None
    
    def test_config_section_properties(self):
        """Test configuration section property accessors."""
        config_manager = ConfigManager()
        
        assert isinstance(config_manager.server, ServerConfig)
        assert isinstance(config_manager.word, WordConfig)
        assert isinstance(config_manager.logging, LoggingConfig)
        assert isinstance(config_manager.security, SecurityConfig)
    
    def test_get_config_path(self):
        """Test getting current configuration file path."""
        config_manager = ConfigManager(self.config_path)
        assert config_manager.get_config_path() == self.config_path
        
        config_manager = ConfigManager()
        assert config_manager.get_config_path() is None
    
    def test_reload_config(self):
        """Test reloading configuration from file."""
        # Create initial config
        config_data = {"server": {"port": 8080}}
        with open(self.config_path, 'w') as f:
            yaml.dump(config_data, f)
        
        config_manager = ConfigManager(self.config_path)
        config = config_manager.load_config()
        assert config.server.port == 8080
        
        # Modify config file
        config_data["server"]["port"] = 9090
        with open(self.config_path, 'w') as f:
            yaml.dump(config_data, f)
        
        # Reload and verify changes
        reloaded_config = config_manager.reload_config()
        assert reloaded_config.server.port == 9090
    
    def test_validate_config_valid(self):
        """Test validating valid configuration data."""
        config_manager = ConfigManager()
        config_data = {
            "server": {"port": 8080},
            "word": {"auto_launch": True}
        }
        
        is_valid, errors = config_manager.validate_config(config_data)
        assert is_valid is True
        assert errors == []
    
    def test_validate_config_invalid(self):
        """Test validating invalid configuration data."""
        config_manager = ConfigManager()
        config_data = {
            "server": {"port": 70000},  # Invalid port
            "logging": {"level": "INVALID"}  # Invalid level
        }
        
        is_valid, errors = config_manager.validate_config(config_data)
        assert is_valid is False
        assert len(errors) == 2
        assert any("server.port" in error for error in errors)
        assert any("logging.level" in error for error in errors)
    
    def test_get_config_schema(self):
        """Test getting configuration schema."""
        config_manager = ConfigManager()
        schema = config_manager.get_config_schema()
        
        assert isinstance(schema, dict)
        assert "properties" in schema
        assert "server" in schema["properties"]
        assert "word" in schema["properties"]
        assert "logging" in schema["properties"]
        assert "security" in schema["properties"]
    
    def test_merge_configs_simple(self):
        """Test merging simple configuration dictionaries."""
        config_manager = ConfigManager()
        
        base_config = {
            "server": {"host": "localhost", "port": 8080},
            "word": {"auto_launch": True}
        }
        
        override_config = {
            "server": {"port": 9090},
            "logging": {"level": "DEBUG"}
        }
        
        merged = config_manager.merge_configs(base_config, override_config)
        
        assert merged["server"]["host"] == "localhost"  # From base
        assert merged["server"]["port"] == 9090  # Overridden
        assert merged["word"]["auto_launch"] is True  # From base
        assert merged["logging"]["level"] == "DEBUG"  # From override
    
    def test_merge_configs_nested(self):
        """Test merging nested configuration dictionaries."""
        config_manager = ConfigManager()
        
        base_config = {
            "server": {
                "host": "localhost",
                "port": 8080,
                "settings": {"timeout": 30, "retries": 3}
            }
        }
        
        override_config = {
            "server": {
                "port": 9090,
                "settings": {"timeout": 60}
            }
        }
        
        merged = config_manager.merge_configs(base_config, override_config)
        
        assert merged["server"]["host"] == "localhost"
        assert merged["server"]["port"] == 9090
        assert merged["server"]["settings"]["timeout"] == 60
        assert merged["server"]["settings"]["retries"] == 3
    
    def test_default_config_paths_search(self):
        """Test searching for config in default paths."""
        # Create config in one of the default paths
        default_config_path = "config.yaml"
        config_data = {"server": {"port": 7777}}
        
        try:
            with open(default_config_path, 'w') as f:
                yaml.dump(config_data, f)
            
            config_manager = ConfigManager()
            config = config_manager.load_config()
            
            assert config.server.port == 7777
        finally:
            # Clean up
            if os.path.exists(default_config_path):
                os.remove(default_config_path)


class TestConfigModels:
    """Test configuration model validation."""
    
    def test_server_config_validation(self):
        """Test server configuration validation."""
        # Valid config
        config = ServerConfig(host="localhost", port=8080)
        assert config.host == "localhost"
        assert config.port == 8080
        
        # Invalid port
        with pytest.raises(ValidationError, match="Port must be between 1 and 65535"):
            ServerConfig(port=70000)
        
        with pytest.raises(ValidationError, match="Port must be between 1 and 65535"):
            ServerConfig(port=0)
        
        # Invalid max_concurrent_docs
        with pytest.raises(ValidationError, match="max_concurrent_docs must be at least 1"):
            ServerConfig(max_concurrent_docs=0)
        
        # Invalid timeout
        with pytest.raises(ValidationError, match="timeout_seconds must be at least 1"):
            ServerConfig(timeout_seconds=0)
    
    def test_word_config_validation(self):
        """Test Word configuration validation."""
        # Valid config
        config = WordConfig(auto_launch=True, visible=False)
        assert config.auto_launch is True
        assert config.visible is False
        
        # Valid absolute backup directory
        config = WordConfig(backup_directory="/tmp/backups")
        assert config.backup_directory == "/tmp/backups"
        
        # Invalid relative backup directory
        with pytest.raises(ValidationError, match="backup_directory must be an absolute path"):
            WordConfig(backup_directory="relative/path")
    
    def test_logging_config_validation(self):
        """Test logging configuration validation."""
        # Valid config
        config = LoggingConfig(level="DEBUG", max_size_mb=50)
        assert config.level == "DEBUG"
        assert config.max_size_mb == 50
        
        # Invalid log level
        with pytest.raises(ValidationError, match="level must be one of"):
            LoggingConfig(level="INVALID")
        
        # Invalid max_size_mb
        with pytest.raises(ValidationError, match="max_size_mb must be at least 1"):
            LoggingConfig(max_size_mb=0)
        
        # Invalid backup_count
        with pytest.raises(ValidationError, match="backup_count must be non-negative"):
            LoggingConfig(backup_count=-1)
    
    def test_security_config_validation(self):
        """Test security configuration validation."""
        # Valid config
        config = SecurityConfig(
            allowed_paths=["~/Documents", "~/Desktop"],
            max_file_size_mb=100
        )
        assert len(config.allowed_paths) == 2
        assert config.max_file_size_mb == 100
        
        # Empty allowed_paths
        with pytest.raises(ValidationError, match="allowed_paths cannot be empty"):
            SecurityConfig(allowed_paths=[])
        
        # Invalid max_file_size_mb
        with pytest.raises(ValidationError, match="max_file_size_mb must be at least 1"):
            SecurityConfig(max_file_size_mb=0)
    
    def test_app_config_validation(self):
        """Test main application configuration validation."""
        # Valid config with defaults
        config = AppConfig()
        assert isinstance(config.server, ServerConfig)
        assert isinstance(config.word, WordConfig)
        assert isinstance(config.logging, LoggingConfig)
        assert isinstance(config.security, SecurityConfig)
        
        # Valid config with custom values
        config = AppConfig(
            server={"port": 9090},
            word={"auto_launch": False},
            logging={"level": "DEBUG"},
            security={"max_file_size_mb": 200}
        )
        assert config.server.port == 9090
        assert config.word.auto_launch is False
        assert config.logging.level == "DEBUG"
        assert config.security.max_file_size_mb == 200
        
        # Invalid nested config
        with pytest.raises(ValidationError):
            AppConfig(server={"port": 70000})
    
    def test_config_extra_fields_forbidden(self):
        """Test that extra fields are forbidden in configuration."""
        with pytest.raises(ValidationError, match="Extra inputs are not permitted"):
            AppConfig(unknown_field="value")
        
        with pytest.raises(ValidationError, match="Extra inputs are not permitted"):
            ServerConfig(unknown_field="value")