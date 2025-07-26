"""
Configuration manager for Word MCP Server.
"""

import os
import yaml
from pathlib import Path
from typing import Optional, Dict, Any
from pydantic import ValidationError

from .models import AppConfig, ServerConfig, WordConfig, LoggingConfig, SecurityConfig


class ConfigManager:
    """Manages configuration loading, validation, and access."""
    
    DEFAULT_CONFIG_PATHS = [
        "config.yaml",
        "config.yml", 
        "word_mcp_server.yaml",
        "word_mcp_server.yml",
        os.path.expanduser("~/.word_mcp_server/config.yaml"),
        os.path.expanduser("~/.config/word_mcp_server/config.yaml"),
    ]
    
    def __init__(self, config_path: Optional[str] = None):
        """Initialize configuration manager.
        
        Args:
            config_path: Optional path to configuration file
        """
        self._config: Optional[AppConfig] = None
        self._config_path: Optional[str] = config_path
        
    def load_config(self, config_path: Optional[str] = None) -> AppConfig:
        """Load configuration from file or use defaults.
        
        Args:
            config_path: Optional path to configuration file
            
        Returns:
            Loaded and validated configuration
            
        Raises:
            FileNotFoundError: If specified config file doesn't exist
            ValidationError: If configuration is invalid
            yaml.YAMLError: If YAML parsing fails
        """
        if config_path:
            self._config_path = config_path
            
        config_data = self._load_config_file()
        
        try:
            self._config = AppConfig(**config_data)
            return self._config
        except ValidationError as e:
            # Provide more helpful error messages
            error_details = []
            for error in e.errors():
                field = ".".join(str(loc) for loc in error['loc'])
                message = error['msg']
                error_details.append(f"  - {field}: {message}")
            
            helpful_message = (
                f"Configuration validation failed:\n"
                f"{''.join(error_details)}\n\n"
                f"Please check your configuration file at: {self._config_path or 'default locations'}\n"
                f"You can create a default configuration using: python -m word_mcp_server --create-config"
            )
            raise ValueError(helpful_message)
    
    def _load_config_file(self) -> Dict[str, Any]:
        """Load configuration data from file.
        
        Returns:
            Configuration data dictionary
        """
        if self._config_path:
            # Use specified config path
            if not os.path.exists(self._config_path):
                raise FileNotFoundError(f"Configuration file not found: {self._config_path}")
            return self._load_yaml_file(self._config_path)
        
        # Try default config paths
        for config_path in self.DEFAULT_CONFIG_PATHS:
            if os.path.exists(config_path):
                return self._load_yaml_file(config_path)
        
        # No config file found, use defaults
        return {}
    
    def _load_yaml_file(self, file_path: str) -> Dict[str, Any]:
        """Load YAML file and return data.
        
        Args:
            file_path: Path to YAML file
            
        Returns:
            Parsed YAML data
            
        Raises:
            yaml.YAMLError: If YAML parsing fails
        """
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = yaml.safe_load(f) or {}
                return data
        except yaml.YAMLError as e:
            raise yaml.YAMLError(f"Failed to parse YAML file {file_path}: {e}")
    
    def save_config(self, config_path: Optional[str] = None) -> None:
        """Save current configuration to file.
        
        Args:
            config_path: Optional path to save configuration
        """
        if not self._config:
            raise ValueError("No configuration loaded to save")
            
        save_path = config_path or self._config_path or "config.yaml"
        
        # Ensure directory exists
        Path(save_path).parent.mkdir(parents=True, exist_ok=True)
        
        # Convert config to dict and save as YAML
        config_dict = self._config.model_dump()
        
        try:
            with open(save_path, 'w', encoding='utf-8') as f:
                yaml.dump(config_dict, f, default_flow_style=False, indent=2)
        except Exception as e:
            raise IOError(f"Failed to save configuration to {save_path}: {e}")
    
    def create_default_config(self, config_path: str) -> AppConfig:
        """Create a default configuration file.
        
        Args:
            config_path: Path where to create the config file
            
        Returns:
            Default configuration
        """
        default_config = AppConfig()
        
        # Ensure directory exists
        Path(config_path).parent.mkdir(parents=True, exist_ok=True)
        
        # Save default config
        config_dict = default_config.model_dump()
        
        try:
            with open(config_path, 'w', encoding='utf-8') as f:
                yaml.dump(config_dict, f, default_flow_style=False, indent=2)
        except Exception as e:
            raise IOError(f"Failed to create default configuration at {config_path}: {e}")
        
        self._config = default_config
        self._config_path = config_path
        return default_config
    
    @property
    def config(self) -> AppConfig:
        """Get current configuration.
        
        Returns:
            Current configuration
            
        Raises:
            ValueError: If no configuration is loaded
        """
        if not self._config:
            # Try to load default configuration
            self._config = self.load_config()
        return self._config
    
    @property
    def server(self) -> ServerConfig:
        """Get server configuration."""
        return self.config.server
    
    @property
    def word(self) -> WordConfig:
        """Get Word configuration."""
        return self.config.word
    
    @property
    def logging(self) -> LoggingConfig:
        """Get logging configuration."""
        return self.config.logging
    
    @property
    def security(self) -> SecurityConfig:
        """Get security configuration."""
        return self.config.security
    
    def get_config_path(self) -> Optional[str]:
        """Get current configuration file path."""
        return self._config_path
    
    def reload_config(self) -> AppConfig:
        """Reload configuration from file.
        
        Returns:
            Reloaded configuration
        """
        return self.load_config(self._config_path)
    
    def validate_config(self, config_data: Dict[str, Any]) -> tuple[bool, list[str]]:
        """Validate configuration data without loading it.
        
        Args:
            config_data: Configuration data to validate
            
        Returns:
            Tuple of (is_valid, error_messages)
        """
        try:
            AppConfig(**config_data)
            return True, []
        except ValidationError as e:
            error_messages = []
            for error in e.errors():
                field = ".".join(str(loc) for loc in error['loc'])
                message = error['msg']
                error_messages.append(f"{field}: {message}")
            return False, error_messages
    
    def get_config_schema(self) -> Dict[str, Any]:
        """Get the configuration schema for documentation.
        
        Returns:
            Configuration schema dictionary
        """
        return AppConfig.model_json_schema()
    
    def merge_configs(self, base_config: Dict[str, Any], override_config: Dict[str, Any]) -> Dict[str, Any]:
        """Merge two configuration dictionaries.
        
        Args:
            base_config: Base configuration
            override_config: Configuration to override with
            
        Returns:
            Merged configuration
        """
        merged = base_config.copy()
        
        for key, value in override_config.items():
            if key in merged and isinstance(merged[key], dict) and isinstance(value, dict):
                merged[key] = self.merge_configs(merged[key], value)
            else:
                merged[key] = value
        
        return merged