"""
Utility functions and helpers for Word MCP Server.
"""

from .errors import ConnectionError, DocumentError, OperationError, WordMCPError
from .logging import setup_logging
from .validators import validate_document_id, validate_path

__all__ = [
    "WordMCPError",
    "ConnectionError",
    "DocumentError",
    "OperationError",
    "setup_logging",
    "validate_path",
    "validate_document_id",
]
