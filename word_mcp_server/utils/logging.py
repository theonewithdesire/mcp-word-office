"""
Logging utilities for Word MCP Server.
"""

import logging
import logging.handlers
import json
import sys
import traceback
from pathlib import Path
from typing import Optional, Dict, Any
from datetime import datetime
from contextlib import contextmanager

from ..config.models import LoggingConfig


class StructuredFormatter(logging.Formatter):
    """Custom formatter that outputs structured JSON logs."""
    
    def format(self, record: logging.LogRecord) -> str:
        """Format log record as structured JSON."""
        log_entry = {
            "timestamp": datetime.fromtimestamp(record.created).isoformat(),
            "level": record.levelname,
            "logger": record.name,
            "message": record.getMessage(),
            "module": record.module,
            "function": record.funcName,
            "line": record.lineno
        }
        
        # Add extra fields if present
        if hasattr(record, 'error_code'):
            log_entry["error_code"] = record.error_code
        if hasattr(record, 'details'):
            log_entry["details"] = record.details
        if hasattr(record, 'context'):
            log_entry["context"] = record.context
        if hasattr(record, 'operation'):
            log_entry["operation"] = record.operation
        if hasattr(record, 'doc_id'):
            log_entry["doc_id"] = record.doc_id
        if hasattr(record, 'duration'):
            log_entry["duration_ms"] = record.duration
        
        # Add exception info if present
        if record.exc_info:
            log_entry["exception"] = {
                "type": record.exc_info[0].__name__,
                "message": str(record.exc_info[1]),
                "traceback": traceback.format_exception(*record.exc_info)
            }
        
        return json.dumps(log_entry, ensure_ascii=False)


class PerformanceLogger:
    """Logger for tracking operation performance."""
    
    def __init__(self, logger: logging.Logger):
        self.logger = logger
        self.operation_times = {}
    
    @contextmanager
    def log_operation(self, operation_name: str, context: Dict[str, Any] = None):
        """Context manager to log operation duration."""
        start_time = datetime.now()
        context = context or {}
        
        self.logger.info(
            f"Starting operation: {operation_name}",
            extra={"operation": operation_name, "context": context}
        )
        
        try:
            yield
            duration = (datetime.now() - start_time).total_seconds() * 1000
            
            self.logger.info(
                f"Completed operation: {operation_name}",
                extra={
                    "operation": operation_name,
                    "duration": duration,
                    "context": context
                }
            )
            
            # Track operation times for statistics
            if operation_name not in self.operation_times:
                self.operation_times[operation_name] = []
            self.operation_times[operation_name].append(duration)
            
        except Exception as e:
            duration = (datetime.now() - start_time).total_seconds() * 1000
            
            self.logger.error(
                f"Failed operation: {operation_name}",
                extra={
                    "operation": operation_name,
                    "duration": duration,
                    "context": context,
                    "error": str(e)
                },
                exc_info=True
            )
            raise
    
    def get_performance_stats(self) -> Dict[str, Dict[str, float]]:
        """Get performance statistics for all operations."""
        stats = {}
        for operation, times in self.operation_times.items():
            if times:
                stats[operation] = {
                    "count": len(times),
                    "avg_ms": sum(times) / len(times),
                    "min_ms": min(times),
                    "max_ms": max(times),
                    "total_ms": sum(times)
                }
        return stats
    
    def reset_performance_stats(self):
        """Reset performance statistics."""
        self.operation_times.clear()


class SecurityLogger:
    """Logger for security-related events."""
    
    def __init__(self, logger: logging.Logger):
        self.logger = logger
    
    def log_access_attempt(self, path: str, operation: str, success: bool, user_context: Dict[str, Any] = None):
        """Log file access attempts."""
        level = logging.INFO if success else logging.WARNING
        message = f"File access {'granted' if success else 'denied'}: {operation} on {path}"
        
        self.logger.log(
            level,
            message,
            extra={
                "security_event": "file_access",
                "path": path,
                "operation": operation,
                "success": success,
                "user_context": user_context or {}
            }
        )
    
    def log_permission_check(self, resource: str, permission: str, granted: bool, reason: str = None):
        """Log permission checks."""
        level = logging.INFO if granted else logging.WARNING
        message = f"Permission {'granted' if granted else 'denied'}: {permission} on {resource}"
        
        extra = {
            "security_event": "permission_check",
            "resource": resource,
            "permission": permission,
            "granted": granted
        }
        
        if reason:
            extra["reason"] = reason
        
        self.logger.log(level, message, extra=extra)
    
    def log_configuration_change(self, setting: str, old_value: Any, new_value: Any, user_context: Dict[str, Any] = None):
        """Log configuration changes."""
        self.logger.warning(
            f"Configuration changed: {setting}",
            extra={
                "security_event": "config_change",
                "setting": setting,
                "old_value": str(old_value),
                "new_value": str(new_value),
                "user_context": user_context or {}
            }
        )


class AuditLogger:
    """Logger for audit trail of operations."""
    
    def __init__(self, logger: logging.Logger):
        self.logger = logger
    
    def log_document_operation(self, operation: str, doc_id: str, path: str = None, 
                             success: bool = True, details: Dict[str, Any] = None):
        """Log document operations for audit trail."""
        message = f"Document {operation}: {doc_id}"
        if path:
            message += f" ({path})"
        
        extra = {
            "audit_event": "document_operation",
            "operation": operation,
            "doc_id": doc_id,
            "success": success
        }
        
        if path:
            extra["path"] = path
        if details:
            extra["details"] = details
        
        level = logging.INFO if success else logging.ERROR
        self.logger.log(level, message, extra=extra)
    
    def log_tool_call(self, tool_name: str, arguments: Dict[str, Any], success: bool = True, 
                     result: Any = None, error: str = None):
        """Log MCP tool calls for audit trail."""
        message = f"Tool call: {tool_name}"
        
        extra = {
            "audit_event": "tool_call",
            "tool_name": tool_name,
            "arguments": arguments,
            "success": success
        }
        
        if result is not None:
            extra["result"] = str(result)[:1000]  # Truncate long results
        if error:
            extra["error"] = error
        
        level = logging.INFO if success else logging.ERROR
        self.logger.log(level, message, extra=extra)


def setup_logging(config: LoggingConfig) -> logging.Logger:
    """Setup comprehensive logging configuration.
    
    Args:
        config: Logging configuration
        
    Returns:
        Configured logger with enhanced features
    """
    # Create main logger
    logger = logging.getLogger("word_mcp_server")
    logger.setLevel(getattr(logging, config.level))
    
    # Clear existing handlers to avoid duplicates
    logger.handlers.clear()
    
    # Create formatters
    console_formatter = logging.Formatter(config.format)
    
    # Console handler - use stderr for MCP compatibility (stdout is reserved for JSON-RPC)
    if config.console:
        console_handler = logging.StreamHandler(sys.stderr)
        console_handler.setLevel(getattr(logging, config.level))
        console_handler.setFormatter(console_formatter)
        
        # Add color support for console output
        if hasattr(console_handler.stream, 'isatty') and console_handler.stream.isatty():
            console_handler.setFormatter(ColoredFormatter(config.format))
        
        logger.addHandler(console_handler)
    
    # File handler with rotation (if specified)
    if config.file:
        # Ensure log directory exists
        log_path = Path(config.file)
        log_path.parent.mkdir(parents=True, exist_ok=True)
        
        # Main log file with rotation
        file_handler = logging.handlers.RotatingFileHandler(
            config.file,
            maxBytes=config.max_size_mb * 1024 * 1024,  # Convert MB to bytes
            backupCount=config.backup_count,
            encoding='utf-8'
        )
        file_handler.setLevel(getattr(logging, config.level))
        file_handler.setFormatter(console_formatter)
        logger.addHandler(file_handler)
        
        # Structured JSON log file for machine processing
        json_log_path = log_path.with_suffix('.json.log')
        json_handler = logging.handlers.RotatingFileHandler(
            json_log_path,
            maxBytes=config.max_size_mb * 1024 * 1024,
            backupCount=config.backup_count,
            encoding='utf-8'
        )
        json_handler.setLevel(getattr(logging, config.level))
        json_handler.setFormatter(StructuredFormatter())
        logger.addHandler(json_handler)
        
        # Error-only log file
        error_log_path = log_path.with_suffix('.error.log')
        error_handler = logging.handlers.RotatingFileHandler(
            error_log_path,
            maxBytes=config.max_size_mb * 1024 * 1024,
            backupCount=config.backup_count,
            encoding='utf-8'
        )
        error_handler.setLevel(logging.ERROR)
        error_handler.setFormatter(console_formatter)
        logger.addHandler(error_handler)
    
    # Set up root logger to capture all logs
    root_logger = logging.getLogger()
    if not root_logger.handlers:
        root_logger.setLevel(logging.WARNING)
        root_handler = logging.StreamHandler()
        root_handler.setFormatter(logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        ))
        root_logger.addHandler(root_handler)
    
    logger.info(f"Logging initialized with level {config.level}")
    return logger


class ColoredFormatter(logging.Formatter):
    """Formatter that adds colors to console output."""
    
    COLORS = {
        'DEBUG': '\033[36m',      # Cyan
        'INFO': '\033[32m',       # Green
        'WARNING': '\033[33m',    # Yellow
        'ERROR': '\033[31m',      # Red
        'CRITICAL': '\033[35m',   # Magenta
        'RESET': '\033[0m'        # Reset
    }
    
    def format(self, record):
        log_color = self.COLORS.get(record.levelname, self.COLORS['RESET'])
        record.levelname = f"{log_color}{record.levelname}{self.COLORS['RESET']}"
        return super().format(record)


def get_logger(name: Optional[str] = None) -> logging.Logger:
    """Get a logger instance.
    
    Args:
        name: Logger name (defaults to word_mcp_server)
        
    Returns:
        Logger instance
    """
    if name:
        return logging.getLogger(f"word_mcp_server.{name}")
    return logging.getLogger("word_mcp_server")


def get_performance_logger(name: Optional[str] = None) -> PerformanceLogger:
    """Get a performance logger instance.
    
    Args:
        name: Logger name
        
    Returns:
        PerformanceLogger instance
    """
    logger = get_logger(name)
    return PerformanceLogger(logger)


def get_security_logger(name: Optional[str] = None) -> SecurityLogger:
    """Get a security logger instance.
    
    Args:
        name: Logger name
        
    Returns:
        SecurityLogger instance
    """
    logger = get_logger(name)
    return SecurityLogger(logger)


def get_audit_logger(name: Optional[str] = None) -> AuditLogger:
    """Get an audit logger instance.
    
    Args:
        name: Logger name
        
    Returns:
        AuditLogger instance
    """
    logger = get_logger(name)
    return AuditLogger(logger)


def configure_graceful_degradation():
    """Configure logging to handle graceful degradation scenarios."""
    # Set up a fallback handler that writes to stderr if all else fails
    fallback_handler = logging.StreamHandler(sys.stderr)
    fallback_handler.setLevel(logging.ERROR)
    fallback_handler.setFormatter(logging.Formatter(
        'FALLBACK: %(asctime)s - %(levelname)s - %(message)s'
    ))
    
    # Add to root logger as last resort
    root_logger = logging.getLogger()
    root_logger.addHandler(fallback_handler)
    
    # Configure to not propagate to avoid duplicate messages
    word_logger = logging.getLogger("word_mcp_server")
    word_logger.propagate = False