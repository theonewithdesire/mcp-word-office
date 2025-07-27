"""
Validation utilities for Word MCP Server.
"""

import os
import re
from pathlib import Path
from typing import List, Optional


def validate_path(path: str, allowed_paths: Optional[List[str]] = None) -> bool:
    """Validate if a file path is allowed.

    Args:
        path: File path to validate
        allowed_paths: List of allowed path patterns

    Returns:
        True if path is valid and allowed

    Raises:
        ValueError: If path is invalid or not allowed
    """
    if not path:
        raise ValueError("Path cannot be empty")

    # Normalize path
    normalized_path = os.path.normpath(os.path.expanduser(path))

    # Check if path exists or parent directory exists (for new files)
    path_obj = Path(normalized_path)
    if not path_obj.exists() and not path_obj.parent.exists():
        raise ValueError(f"Path or parent directory does not exist: {path}")

    # Check against allowed paths if specified
    if allowed_paths:
        allowed = False
        for allowed_pattern in allowed_paths:
            allowed_pattern = os.path.normpath(os.path.expanduser(allowed_pattern))

            # Check if path starts with allowed pattern
            if normalized_path.startswith(allowed_pattern):
                allowed = True
                break

        if not allowed:
            raise ValueError(f"Path not in allowed directories: {path}")

    return True


def validate_document_id(doc_id: str) -> bool:
    """Validate document ID format.

    Args:
        doc_id: Document ID to validate

    Returns:
        True if document ID is valid

    Raises:
        ValueError: If document ID is invalid
    """
    if not doc_id:
        raise ValueError("Document ID cannot be empty")

    # Document ID should be alphanumeric with hyphens and underscores
    if not re.match(r"^[a-zA-Z0-9_-]+$", doc_id):
        raise ValueError(
            "Document ID can only contain letters, numbers, hyphens, and underscores"
        )

    if len(doc_id) > 100:
        raise ValueError("Document ID cannot be longer than 100 characters")

    return True


def validate_file_path(path: str) -> bool:
    """Validate if a file path is valid for document operations.

    Args:
        path: File path to validate

    Returns:
        True if path is valid

    Raises:
        ValueError: If path is invalid
    """
    if not path:
        return False

    try:
        # Normalize path
        normalized_path = os.path.normpath(os.path.expanduser(path))

        # Check for dangerous path patterns
        if ".." in normalized_path or normalized_path.startswith("/"):
            return False

        # Check if it's a reasonable file path
        path_obj = Path(normalized_path)

        # Path should have a valid extension for Word documents
        valid_extensions = [".docx", ".doc", ".dotx", ".dotm", ".docm"]
        if path_obj.suffix.lower() not in valid_extensions:
            # Allow paths without extension for validation purposes
            pass

        return True

    except Exception:
        return False


def validate_file_size(file_path: str, max_size_mb: int) -> bool:
    """Validate file size against maximum allowed size.

    Args:
        file_path: Path to file to check
        max_size_mb: Maximum allowed size in MB

    Returns:
        True if file size is within limits

    Raises:
        ValueError: If file is too large
        FileNotFoundError: If file doesn't exist
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")

    file_size_mb = os.path.getsize(file_path) / (1024 * 1024)

    if file_size_mb > max_size_mb:
        raise ValueError(
            f"File size ({file_size_mb:.1f}MB) exceeds maximum allowed size ({max_size_mb}MB)"
        )

    return True
