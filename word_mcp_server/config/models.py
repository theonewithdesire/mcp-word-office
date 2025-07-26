"""
Configuration data models for Word MCP Server.
"""

from typing import List, Optional
from pydantic import BaseModel, Field, field_validator, ConfigDict
from pathlib import Path


class ServerConfig(BaseModel):
    """Server configuration settings."""
    
    model_config = ConfigDict(extra="forbid")
    
    host: str = Field(default="localhost", description="Server host address")
    port: int = Field(default=8080, description="Server port number")
    max_concurrent_docs: int = Field(default=10, description="Maximum concurrent documents")
    timeout_seconds: int = Field(default=30, description="Operation timeout in seconds")
    
    @field_validator('port')
    @classmethod
    def validate_port(cls, v):
        if not 1 <= v <= 65535:
            raise ValueError('Port must be between 1 and 65535')
        return v
    
    @field_validator('max_concurrent_docs')
    @classmethod
    def validate_max_concurrent_docs(cls, v):
        if v < 1:
            raise ValueError('max_concurrent_docs must be at least 1')
        return v
    
    @field_validator('timeout_seconds')
    @classmethod
    def validate_timeout(cls, v):
        if v < 1:
            raise ValueError('timeout_seconds must be at least 1')
        return v


class WordConfig(BaseModel):
    """Microsoft Word application configuration."""
    
    model_config = ConfigDict(extra="forbid")
    
    auto_launch: bool = Field(default=True, description="Auto-launch Word if not running")
    visible: bool = Field(default=False, description="Make Word application visible")
    save_on_exit: bool = Field(default=True, description="Save documents on exit")
    backup_enabled: bool = Field(default=True, description="Enable document backups")
    backup_directory: Optional[str] = Field(default=None, description="Backup directory path")
    
    @field_validator('backup_directory')
    @classmethod
    def validate_backup_directory(cls, v):
        if v is not None:
            path = Path(v)
            if not path.is_absolute():
                raise ValueError('backup_directory must be an absolute path')
        return v


class LoggingConfig(BaseModel):
    """Logging configuration settings."""
    
    model_config = ConfigDict(extra="forbid")
    
    level: str = Field(default="INFO", description="Logging level")
    file: Optional[str] = Field(default="word_mcp_server.log", description="Log file path")
    max_size_mb: int = Field(default=100, description="Maximum log file size in MB")
    backup_count: int = Field(default=5, description="Number of backup log files")
    console: bool = Field(default=False, description="Enable console logging (disabled for MCP compatibility)")
    format: str = Field(
        default="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
        description="Log message format"
    )
    
    @field_validator('level')
    @classmethod
    def validate_level(cls, v):
        valid_levels = ['DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITICAL']
        if v.upper() not in valid_levels:
            raise ValueError(f'level must be one of {valid_levels}')
        return v.upper()
    
    @field_validator('max_size_mb')
    @classmethod
    def validate_max_size(cls, v):
        if v < 1:
            raise ValueError('max_size_mb must be at least 1')
        return v
    
    @field_validator('backup_count')
    @classmethod
    def validate_backup_count(cls, v):
        if v < 0:
            raise ValueError('backup_count must be non-negative')
        return v


class SecurityConfig(BaseModel):
    """Security configuration settings."""
    
    model_config = ConfigDict(extra="forbid")
    
    allowed_paths: List[str] = Field(
        default_factory=lambda: ["~/Documents", "~/Desktop"],
        description="Allowed file system paths"
    )
    enable_macros: bool = Field(default=False, description="Enable macro execution")
    max_file_size_mb: int = Field(default=50, description="Maximum file size in MB")
    require_confirmation: bool = Field(
        default=True, 
        description="Require confirmation for destructive operations"
    )
    
    @field_validator('allowed_paths')
    @classmethod
    def validate_allowed_paths(cls, v):
        if not v:
            raise ValueError('allowed_paths cannot be empty')
        return v
    
    @field_validator('max_file_size_mb')
    @classmethod
    def validate_max_file_size(cls, v):
        if v < 1:
            raise ValueError('max_file_size_mb must be at least 1')
        return v


class AppConfig(BaseModel):
    """Main application configuration."""
    
    server: ServerConfig = Field(default_factory=ServerConfig)
    word: WordConfig = Field(default_factory=WordConfig)
    logging: LoggingConfig = Field(default_factory=LoggingConfig)
    security: SecurityConfig = Field(default_factory=SecurityConfig)
    
    model_config = ConfigDict(
        validate_assignment=True,
        extra="forbid"
    )