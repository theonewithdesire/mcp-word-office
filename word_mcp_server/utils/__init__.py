"""
Utility functions and helpers for Word MCP Server.
"""

from .errors import WordMCPError, ConnectionError, DocumentError, OperationError
from .logging import setup_logging
from .validators import validate_path, validate_document_id

__all__ = [
    "WordMCPError", 
    "ConnectionError", 
    "DocumentError", 
    "OperationError",
    "setup_logging",
    "validate_path",
    "validate_document_id"
]