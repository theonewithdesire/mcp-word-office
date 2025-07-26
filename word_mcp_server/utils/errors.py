"""
Custom exception classes for Word MCP Server.
"""

from enum import Enum
from typing import Dict, List, Optional, Any
import logging


class ErrorCode(Enum):
    """Error codes for Word MCP Server operations."""
    
    # Connection errors
    WORD_CONNECTION_FAILED = "WORD_CONNECTION_FAILED"
    WORD_NOT_INSTALLED = "WORD_NOT_INSTALLED"
    COM_INTERFACE_ERROR = "COM_INTERFACE_ERROR"
    WORD_CRASHED = "WORD_CRASHED"
    WORD_UNRESPONSIVE = "WORD_UNRESPONSIVE"
    
    # Document errors
    DOCUMENT_NOT_FOUND = "DOCUMENT_NOT_FOUND"
    DOCUMENT_ACCESS_DENIED = "DOCUMENT_ACCESS_DENIED"
    DOCUMENT_CORRUPTED = "DOCUMENT_CORRUPTED"
    DOCUMENT_LOCKED = "DOCUMENT_LOCKED"
    DOCUMENT_READ_ONLY = "DOCUMENT_READ_ONLY"
    DOCUMENT_TOO_LARGE = "DOCUMENT_TOO_LARGE"
    
    # Operation errors
    INVALID_PARAMETERS = "INVALID_PARAMETERS"
    OPERATION_FAILED = "OPERATION_FAILED"
    UNSUPPORTED_OPERATION = "UNSUPPORTED_OPERATION"
    TIMEOUT_ERROR = "TIMEOUT_ERROR"
    MEMORY_ERROR = "MEMORY_ERROR"
    DISK_SPACE_ERROR = "DISK_SPACE_ERROR"
    
    # MCP protocol errors
    INVALID_TOOL_CALL = "INVALID_TOOL_CALL"
    TOOL_NOT_FOUND = "TOOL_NOT_FOUND"
    PROTOCOL_ERROR = "PROTOCOL_ERROR"
    SERIALIZATION_ERROR = "SERIALIZATION_ERROR"
    
    # General errors
    UNKNOWN_ERROR = "UNKNOWN_ERROR"
    CONFIGURATION_ERROR = "CONFIGURATION_ERROR"
    PERMISSION_DENIED = "PERMISSION_DENIED"
    NETWORK_ERROR = "NETWORK_ERROR"
    RESOURCE_EXHAUSTED = "RESOURCE_EXHAUSTED"


class ErrorSuggestions:
    """Provides troubleshooting suggestions for different error types."""
    
    SUGGESTIONS = {
        ErrorCode.WORD_CONNECTION_FAILED: [
            "Ensure Microsoft Word is installed on your system",
            "Try running the application as administrator",
            "Check if Word is already running and close any dialog boxes",
            "Restart the Word application and try again",
            "Verify COM registration: run 'regsvr32 /i msword.olb' as administrator"
        ],
        ErrorCode.WORD_NOT_INSTALLED: [
            "Install Microsoft Word 2016 or later",
            "Verify Word installation by opening it manually",
            "Check Windows Programs and Features for Office installation",
            "Reinstall Microsoft Office if necessary"
        ],
        ErrorCode.COM_INTERFACE_ERROR: [
            "Install pywin32: pip install pywin32",
            "Run post-install script: python Scripts/pywin32_postinstall.py -install",
            "Check Windows Event Viewer for COM errors",
            "Try running as administrator to fix COM registration issues"
        ],
        ErrorCode.WORD_CRASHED: [
            "Restart Microsoft Word application",
            "Check for Word updates and install them",
            "Disable Word add-ins that might cause instability",
            "Run Word in safe mode: winword /safe",
            "Check system memory and disk space"
        ],
        ErrorCode.WORD_UNRESPONSIVE: [
            "Wait for Word to respond or force-close it",
            "Check Task Manager for hung Word processes",
            "Restart the Word application",
            "Reduce document complexity or size",
            "Check system resources (CPU, memory)"
        ],
        ErrorCode.DOCUMENT_NOT_FOUND: [
            "Verify the file path is correct and accessible",
            "Check if the file has been moved or deleted",
            "Ensure you have read permissions for the directory",
            "Use absolute file paths instead of relative paths"
        ],
        ErrorCode.DOCUMENT_ACCESS_DENIED: [
            "Check file permissions and ensure you have read/write access",
            "Close the document if it's open in another application",
            "Run the application as administrator",
            "Check if the file is on a network drive with access restrictions"
        ],
        ErrorCode.DOCUMENT_CORRUPTED: [
            "Try opening the document in Word to check for corruption",
            "Use Word's built-in repair feature: File > Open > Browse > Tools > Open and Repair",
            "Restore from a backup copy if available",
            "Try opening in Word's safe mode"
        ],
        ErrorCode.DOCUMENT_LOCKED: [
            "Close the document in other applications",
            "Check if another user has the document open",
            "Wait for the lock to be released",
            "Use Task Manager to end Word processes if necessary"
        ],
        ErrorCode.DOCUMENT_READ_ONLY: [
            "Check file properties and remove read-only attribute",
            "Ensure you have write permissions to the file and directory",
            "Save to a different location if the original is protected",
            "Contact your system administrator for write access"
        ],
        ErrorCode.DOCUMENT_TOO_LARGE: [
            "Break the document into smaller sections",
            "Remove unnecessary images or embedded objects",
            "Increase available system memory",
            "Use a more powerful computer for large documents"
        ],
        ErrorCode.INVALID_PARAMETERS: [
            "Check the parameter values and types",
            "Refer to the API documentation for correct parameter format",
            "Ensure required parameters are provided",
            "Validate parameter ranges and constraints"
        ],
        ErrorCode.TIMEOUT_ERROR: [
            "Increase the timeout value in configuration",
            "Check system performance and available resources",
            "Reduce document complexity",
            "Try the operation again with a smaller scope"
        ],
        ErrorCode.MEMORY_ERROR: [
            "Close other applications to free memory",
            "Restart the application to clear memory leaks",
            "Increase system RAM if possible",
            "Process documents in smaller batches"
        ],
        ErrorCode.DISK_SPACE_ERROR: [
            "Free up disk space on the system drive",
            "Clean temporary files and caches",
            "Move large files to external storage",
            "Check available space in the document directory"
        ],
        ErrorCode.CONFIGURATION_ERROR: [
            "Check the configuration file syntax",
            "Verify all required configuration values are set",
            "Restore default configuration if necessary",
            "Check file permissions for configuration files"
        ],
        ErrorCode.PERMISSION_DENIED: [
            "Run the application as administrator",
            "Check file and directory permissions",
            "Ensure the user has necessary privileges",
            "Contact your system administrator"
        ],
        ErrorCode.NETWORK_ERROR: [
            "Check network connectivity",
            "Verify firewall settings",
            "Try accessing the resource directly",
            "Check proxy settings if applicable"
        ]
    }
    
    @classmethod
    def get_suggestions(cls, error_code: ErrorCode) -> List[str]:
        """Get troubleshooting suggestions for an error code."""
        return cls.SUGGESTIONS.get(error_code, [
            "Check the application logs for more details",
            "Try restarting the application",
            "Contact support if the problem persists"
        ])


class WordMCPError(Exception):
    """Base exception class for Word MCP Server."""
    
    def __init__(self, message: str, error_code: str = None, details: str = None, 
                 suggestions: List[str] = None, context: Dict[str, Any] = None):
        super().__init__(message)
        self.message = message
        self.error_code = error_code or ErrorCode.UNKNOWN_ERROR.value
        self.details = details
        self.suggestions = suggestions or []
        self.context = context or {}
        
        # Auto-generate suggestions if not provided
        if not self.suggestions:
            try:
                error_enum = ErrorCode(self.error_code)
                self.suggestions = ErrorSuggestions.get_suggestions(error_enum)
            except ValueError:
                self.suggestions = ErrorSuggestions.get_suggestions(ErrorCode.UNKNOWN_ERROR)
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert error to dictionary format for JSON serialization."""
        return {
            "error": {
                "code": self.error_code,
                "message": self.message,
                "details": self.details,
                "suggestions": self.suggestions,
                "context": self.context
            }
        }
    
    def log_error(self, logger: logging.Logger):
        """Log the error with appropriate level and context."""
        logger.error(
            f"Error {self.error_code}: {self.message}",
            extra={
                "error_code": self.error_code,
                "details": self.details,
                "context": self.context
            },
            exc_info=True
        )


class ConnectionError(WordMCPError):
    """Raised when connection to Word application fails."""
    
    def __init__(self, message: str = "Failed to connect to Word application", **kwargs):
        if 'error_code' not in kwargs:
            kwargs['error_code'] = ErrorCode.WORD_CONNECTION_FAILED.value
        super().__init__(message, **kwargs)


class DocumentError(WordMCPError):
    """Raised when document operations fail."""
    
    def __init__(self, message: str, **kwargs):
        if 'error_code' not in kwargs:
            kwargs['error_code'] = ErrorCode.OPERATION_FAILED.value
        super().__init__(message, **kwargs)


class OperationError(WordMCPError):
    """Raised when Word operations fail."""
    
    def __init__(self, message: str, **kwargs):
        if 'error_code' not in kwargs:
            kwargs['error_code'] = ErrorCode.OPERATION_FAILED.value
        super().__init__(message, **kwargs)


class TimeoutError(WordMCPError):
    """Raised when operations timeout."""
    
    def __init__(self, message: str = "Operation timed out", **kwargs):
        if 'error_code' not in kwargs:
            kwargs['error_code'] = ErrorCode.TIMEOUT_ERROR.value
        super().__init__(message, **kwargs)


class ConfigurationError(WordMCPError):
    """Raised when configuration is invalid."""
    
    def __init__(self, message: str = "Configuration error", **kwargs):
        if 'error_code' not in kwargs:
            kwargs['error_code'] = ErrorCode.CONFIGURATION_ERROR.value
        super().__init__(message, **kwargs)


class ErrorHandler:
    """Centralized error handling and recovery system."""
    
    def __init__(self, logger: logging.Logger):
        self.logger = logger
        self.error_counts = {}
        self.recovery_strategies = {}
    
    def handle_error(self, error: Exception, context: Dict[str, Any] = None) -> Dict[str, Any]:
        """Handle an error with appropriate logging and recovery strategies."""
        context = context or {}
        
        # Convert to WordMCPError if not already
        if not isinstance(error, WordMCPError):
            error = WordMCPError(
                message=str(error),
                error_code=ErrorCode.UNKNOWN_ERROR.value,
                context=context
            )
        
        # Log the error
        error.log_error(self.logger)
        
        # Track error frequency
        self._track_error(error.error_code)
        
        # Attempt recovery if strategy exists
        recovery_result = self._attempt_recovery(error)
        
        # Return error response
        error_response = error.to_dict()
        if recovery_result:
            error_response["recovery"] = recovery_result
        
        return error_response
    
    def _track_error(self, error_code: str):
        """Track error frequency for monitoring."""
        self.error_counts[error_code] = self.error_counts.get(error_code, 0) + 1
        
        # Log warning if error is frequent
        if self.error_counts[error_code] > 5:
            self.logger.warning(
                f"Frequent error detected: {error_code} occurred {self.error_counts[error_code]} times"
            )
    
    def _attempt_recovery(self, error: WordMCPError) -> Optional[Dict[str, Any]]:
        """Attempt to recover from specific error types."""
        recovery_strategy = self.recovery_strategies.get(error.error_code)
        if recovery_strategy:
            try:
                return recovery_strategy(error)
            except Exception as e:
                self.logger.warning(f"Recovery strategy failed for {error.error_code}: {e}")
        return None
    
    def register_recovery_strategy(self, error_code: str, strategy_func):
        """Register a recovery strategy for a specific error code."""
        self.recovery_strategies[error_code] = strategy_func
    
    def get_error_statistics(self) -> Dict[str, int]:
        """Get error frequency statistics."""
        return self.error_counts.copy()
    
    def reset_error_statistics(self):
        """Reset error frequency tracking."""
        self.error_counts.clear()