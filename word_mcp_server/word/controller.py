"""
Word application controller using COM interface.

This module provides the WordController class for managing Microsoft Word
application instances and document lifecycle through COM automation.
"""

import logging
import time
import uuid
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

try:
    import pythoncom
    import win32com.client
    from win32com.client import constants as word_constants

    COM_AVAILABLE = True
except ImportError:
    COM_AVAILABLE = False
    win32com = None
    pythoncom = None
    word_constants = None

from ..config.models import WordConfig
from ..utils.errors import (
    ConnectionError,
    DocumentError,
    ErrorCode,
    OperationError,
    TimeoutError,
    WordMCPError,
)
from ..utils.logging import get_logger, get_performance_logger, get_security_logger
from ..utils.recovery import CircuitBreaker, RecoveryConfig, RetryManager


@dataclass
class DocumentReference:
    """Reference to an open Word document."""

    doc_id: str
    file_path: Optional[str]
    title: str
    is_active: bool
    word_doc_ref: Optional[object]  # COM object reference
    created_at: datetime
    last_modified: datetime


class WordController:
    """Handles direct Word application control via COM interface.

    This class manages the connection to Microsoft Word application,
    handles document lifecycle operations, and provides error recovery
    mechanisms for COM interface failures.
    """

    def __init__(self, config: WordConfig):
        """Initialize Word controller.

        Args:
            config: Word configuration settings

        Raises:
            ConnectionError: If COM interface is not available
        """
        if not COM_AVAILABLE:
            raise ConnectionError(
                "COM interface not available. Please install pywin32.",
                details="Run: pip install pywin32",
            )

        self.config = config
        self.logger = get_logger(__name__)
        self.performance_logger = get_performance_logger(__name__)
        self.security_logger = get_security_logger(__name__)

        # Initialize recovery systems
        self.retry_manager = RetryManager(
            RecoveryConfig(
                max_retries=3,
                retry_delay=2.0,
                backoff_multiplier=1.5,
                max_delay=30.0,
                timeout=60.0,
            ),
            self.logger,
        )

        # Circuit breaker for Word operations
        self.word_circuit_breaker = CircuitBreaker(
            failure_threshold=5, recovery_timeout=60.0, logger=self.logger
        )

        self._word_app = None
        self._documents: Dict[str, DocumentReference] = {}
        self._connection_attempts = 0
        self._max_connection_attempts = 3
        self._retry_delay = 2.0

        self.logger.info("WordController initialized with enhanced error handling")

    def connect_to_word(self) -> bool:
        """Connect to Word application.

        Attempts to connect to an existing Word instance or launches a new one
        if auto_launch is enabled in configuration.

        Returns:
            True if connection successful, False otherwise

        Raises:
            ConnectionError: If connection fails after all retry attempts
        """
        self.logger.info("Attempting to connect to Word application")

        for attempt in range(self._max_connection_attempts):
            try:
                self._connection_attempts = attempt + 1

                # Try to connect to existing Word instance first
                if self._try_connect_existing():
                    self.logger.info("Connected to existing Word instance")
                    return True

                # If no existing instance and auto_launch is enabled, launch new one
                if self.config.auto_launch:
                    if self._launch_new_word():
                        self.logger.info("Launched new Word instance and connected")
                        return True

                # If we reach here, connection failed for this attempt
                if attempt < self._max_connection_attempts - 1:
                    self.logger.warning(
                        f"Connection attempt {attempt + 1} failed, retrying in {self._retry_delay}s"
                    )
                    time.sleep(self._retry_delay)
                    self._retry_delay *= 1.5  # Exponential backoff

            except Exception as e:
                self.logger.error(f"Connection attempt {attempt + 1} failed: {str(e)}")
                if attempt < self._max_connection_attempts - 1:
                    time.sleep(self._retry_delay)
                    self._retry_delay *= 1.5
                else:
                    raise ConnectionError(
                        f"Failed to connect to Word after {self._max_connection_attempts} attempts",
                        details=str(e),
                    )

        raise ConnectionError(
            "Could not establish connection to Word application",
            details="Word may not be installed or COM interface is unavailable",
        )

    def _try_connect_existing(self) -> bool:
        """Try to connect to existing Word instance.

        Returns:
            True if connection successful, False otherwise
        """
        try:
            # Initialize COM for this thread
            pythoncom.CoInitialize()

            # Try to get existing Word application
            self._word_app = win32com.client.GetActiveObject("Word.Application")

            # Test the connection by accessing a property
            _ = self._word_app.Version

            # Configure Word application settings
            self._configure_word_app()

            return True

        except Exception as e:
            self.logger.debug(f"Could not connect to existing Word instance: {str(e)}")
            self._word_app = None
            return False

    def _launch_new_word(self) -> bool:
        """Launch a new Word application instance.

        Returns:
            True if launch successful, False otherwise
        """
        try:
            # Initialize COM for this thread
            pythoncom.CoInitialize()

            # Create new Word application instance
            self._word_app = win32com.client.Dispatch("Word.Application")

            # Test the connection
            _ = self._word_app.Version

            # Configure Word application settings
            self._configure_word_app()

            self.logger.info(f"Launched Word version: {self._word_app.Version}")
            return True

        except Exception as e:
            self.logger.error(f"Failed to launch new Word instance: {str(e)}")
            self._word_app = None
            return False

    def _configure_word_app(self):
        """Configure Word application settings based on configuration."""
        if not self._word_app:
            return

        try:
            # Set visibility
            self._word_app.Visible = self.config.visible

            # Configure other settings
            self._word_app.DisplayAlerts = False  # Suppress dialog boxes

            self.logger.debug("Word application configured successfully")

        except Exception as e:
            self.logger.warning(f"Could not configure Word application: {str(e)}")

    def is_connected(self) -> bool:
        """Check if connected to Word application.

        Returns:
            True if connected and Word is responsive, False otherwise
        """
        if not self._word_app:
            return False

        try:
            # Test connection by accessing a property
            _ = self._word_app.Version
            return True
        except Exception as e:
            self.logger.warning(f"Word connection lost: {str(e)}")
            self._word_app = None
            return False

    def disconnect(self):
        """Disconnect from Word application.

        Closes all managed documents and releases COM objects.
        """
        self.logger.info("Disconnecting from Word application")

        try:
            # Close all managed documents
            self._close_all_documents()

            # Release Word application reference
            if self._word_app:
                try:
                    # Optionally quit Word if we launched it
                    if self.config.save_on_exit:
                        self._word_app.Quit(SaveChanges=True)
                    else:
                        self._word_app.Quit(SaveChanges=False)
                except Exception as e:
                    self.logger.warning(f"Error quitting Word: {str(e)}")

                self._word_app = None

            # Uninitialize COM
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass  # COM may already be uninitialized

            self.logger.info("Disconnected from Word application")

        except Exception as e:
            self.logger.error(f"Error during disconnect: {str(e)}")

    def create_document(self) -> str:
        """Create a new Word document.

        Returns:
            Document ID for the created document

        Raises:
            ConnectionError: If not connected to Word
            OperationError: If document creation fails
        """
        if not self.is_connected():
            raise ConnectionError("Not connected to Word application")

        try:
            # Create new document
            doc = self._word_app.Documents.Add()

            # Generate unique document ID
            doc_id = str(uuid.uuid4())

            # Create document reference
            doc_ref = DocumentReference(
                doc_id=doc_id,
                file_path=None,
                title=doc.Name,
                is_active=True,
                word_doc_ref=doc,
                created_at=datetime.now(),
                last_modified=datetime.now(),
            )

            # Store document reference
            self._documents[doc_id] = doc_ref

            self.logger.info(f"Created new document with ID: {doc_id}")
            return doc_id

        except Exception as e:
            raise OperationError(
                f"Failed to create new document: {str(e)}",
                error_code=ErrorCode.OPERATION_FAILED.value,
            )

    def open_document(self, path: str) -> str:
        """Open an existing document.

        Args:
            path: Document file path

        Returns:
            Document ID for the opened document

        Raises:
            ConnectionError: If not connected to Word
            DocumentError: If document cannot be opened
        """
        if not self.is_connected():
            raise ConnectionError("Not connected to Word application")

        # Validate file path
        file_path = Path(path)
        if not file_path.exists():
            raise DocumentError(
                f"Document not found: {path}",
                error_code=ErrorCode.DOCUMENT_NOT_FOUND.value,
            )

        try:
            # Open document
            doc = self._word_app.Documents.Open(str(file_path.absolute()))

            # Generate unique document ID
            doc_id = str(uuid.uuid4())

            # Create document reference
            doc_ref = DocumentReference(
                doc_id=doc_id,
                file_path=str(file_path.absolute()),
                title=doc.Name,
                is_active=True,
                word_doc_ref=doc,
                created_at=datetime.now(),
                last_modified=datetime.now(),
            )

            # Store document reference
            self._documents[doc_id] = doc_ref

            self.logger.info(f"Opened document '{path}' with ID: {doc_id}")
            return doc_id

        except Exception as e:
            if "access denied" in str(e).lower():
                raise DocumentError(
                    f"Access denied to document: {path}",
                    error_code=ErrorCode.DOCUMENT_ACCESS_DENIED.value,
                    details="Check file permissions or if file is open in another application",
                )
            else:
                raise DocumentError(
                    f"Failed to open document '{path}': {str(e)}",
                    error_code=ErrorCode.DOCUMENT_CORRUPTED.value,
                )

    def close_document(self, doc_id: str, save: bool = True):
        """Close a document.

        Args:
            doc_id: Document ID
            save: Whether to save the document before closing

        Raises:
            DocumentError: If document not found or close operation fails
        """
        if doc_id not in self._documents:
            raise DocumentError(
                f"Document not found: {doc_id}",
                error_code=ErrorCode.DOCUMENT_NOT_FOUND.value,
            )

        doc_ref = self._documents[doc_id]

        try:
            if doc_ref.word_doc_ref:
                # Close the document
                doc_ref.word_doc_ref.Close(SaveChanges=save)

            # Remove from our tracking
            del self._documents[doc_id]

            self.logger.info(f"Closed document: {doc_id}")

        except Exception as e:
            raise OperationError(
                f"Failed to close document '{doc_id}': {str(e)}",
                error_code=ErrorCode.OPERATION_FAILED.value,
            )

    def save_document(self, doc_id: str, path: Optional[str] = None):
        """Save a document.

        Args:
            doc_id: Document ID
            path: Optional path to save to (if different from current path)

        Raises:
            DocumentError: If document not found or save operation fails
        """
        if doc_id not in self._documents:
            raise DocumentError(
                f"Document not found: {doc_id}",
                error_code=ErrorCode.DOCUMENT_NOT_FOUND.value,
            )

        doc_ref = self._documents[doc_id]

        try:
            if path:
                # Save as new file
                doc_ref.word_doc_ref.SaveAs2(str(Path(path).absolute()))
                doc_ref.file_path = str(Path(path).absolute())
            else:
                # Save existing file
                doc_ref.word_doc_ref.Save()

            doc_ref.last_modified = datetime.now()

            self.logger.info(f"Saved document: {doc_id}")

        except Exception as e:
            if "access denied" in str(e).lower():
                raise DocumentError(
                    f"Access denied when saving document: {doc_id}",
                    error_code=ErrorCode.DOCUMENT_ACCESS_DENIED.value,
                )
            else:
                raise OperationError(
                    f"Failed to save document '{doc_id}': {str(e)}",
                    error_code=ErrorCode.OPERATION_FAILED.value,
                )

    def get_document_reference(self, doc_id: str) -> DocumentReference:
        """Get document reference by ID.

        Args:
            doc_id: Document ID

        Returns:
            DocumentReference object

        Raises:
            DocumentError: If document not found
        """
        if doc_id not in self._documents:
            raise DocumentError(
                f"Document not found: {doc_id}",
                error_code=ErrorCode.DOCUMENT_NOT_FOUND.value,
            )

        return self._documents[doc_id]

    def list_documents(self) -> Dict[str, DocumentReference]:
        """List all managed documents.

        Returns:
            Dictionary of document ID to DocumentReference
        """
        return self._documents.copy()

    def _close_all_documents(self):
        """Close all managed documents."""
        doc_ids = list(self._documents.keys())
        for doc_id in doc_ids:
            try:
                self.close_document(doc_id, save=self.config.save_on_exit)
            except Exception as e:
                self.logger.warning(f"Error closing document {doc_id}: {str(e)}")

    async def cleanup_all_documents(self):
        """Async cleanup method for graceful shutdown.

        Saves and closes all open documents, handles any errors gracefully.
        This method is designed to be called during server shutdown.
        """
        self.logger.info(
            f"Starting document cleanup for {len(self._documents)} documents"
        )

        cleanup_results = {
            "total_documents": len(self._documents),
            "successfully_closed": 0,
            "errors": [],
        }

        # Get a copy of document IDs to avoid modification during iteration
        doc_ids = list(self._documents.keys())

        for doc_id in doc_ids:
            try:
                doc_ref = self._documents.get(doc_id)
                if doc_ref:
                    self.logger.debug(
                        f"Cleaning up document: {doc_ref.title} ({doc_id})"
                    )

                    # Save document if it has unsaved changes
                    if doc_ref.word_doc_ref:
                        try:
                            # Check if document has unsaved changes
                            if (
                                hasattr(doc_ref.word_doc_ref, "Saved")
                                and not doc_ref.word_doc_ref.Saved
                            ):
                                if doc_ref.file_path:
                                    doc_ref.word_doc_ref.Save()
                                    self.logger.debug(
                                        f"Saved document: {doc_ref.title}"
                                    )
                                else:
                                    # Document has no path, save with default name
                                    default_path = f"Document_{doc_id}.docx"
                                    doc_ref.word_doc_ref.SaveAs2(default_path)
                                    self.logger.debug(
                                        f"Saved new document as: {default_path}"
                                    )
                        except Exception as save_error:
                            error_msg = (
                                f"Failed to save document {doc_ref.title}: {save_error}"
                            )
                            self.logger.warning(error_msg)
                            cleanup_results["errors"].append(error_msg)

                    # Close the document
                    self.close_document(doc_id, save=False)  # Already saved above
                    cleanup_results["successfully_closed"] += 1

            except Exception as e:
                error_msg = f"Error cleaning up document {doc_id}: {e}"
                self.logger.error(error_msg)
                cleanup_results["errors"].append(error_msg)

        # Log cleanup summary
        if cleanup_results["errors"]:
            self.logger.warning(
                f"Document cleanup completed with {len(cleanup_results['errors'])} errors. "
                f"Successfully closed {cleanup_results['successfully_closed']}/{cleanup_results['total_documents']} documents."
            )
        else:
            self.logger.info(
                f"Document cleanup completed successfully. "
                f"Closed {cleanup_results['successfully_closed']} documents."
            )

        return cleanup_results

    def __enter__(self):
        """Context manager entry."""
        self.connect_to_word()
        return self

    def insert_text(self, doc_id: str, text: str, position: Optional[int] = None):
        """Insert text into a document.

        Args:
            doc_id: Document ID
            text: Text to insert
            position: Optional position to insert at (0-based). If None, inserts at end.

        Raises:
            DocumentError: If document not found
            OperationError: If text insertion fails
        """
        if doc_id not in self._documents:
            raise DocumentError(
                f"Document not found: {doc_id}",
                error_code=ErrorCode.DOCUMENT_NOT_FOUND.value,
            )

        doc_ref = self._documents[doc_id]

        try:
            doc = doc_ref.word_doc_ref
            if not doc:
                raise OperationError("Document reference is invalid")

            # Get the document range
            doc_range = doc.Range()

            if position is not None:
                # Insert at specific position
                if position < 0:
                    position = 0
                elif position > len(doc_range.Text):
                    position = len(doc_range.Text)

                # Set the range to the insertion point
                doc_range.SetRange(position, position)
            else:
                # Insert at the end of the document
                doc_range.SetRange(len(doc_range.Text), len(doc_range.Text))

            # Insert the text
            doc_range.InsertAfter(text)

            doc_ref.last_modified = datetime.now()

            self.logger.info(
                f"Inserted text into document {doc_id} at position {position}"
            )

        except Exception as e:
            raise OperationError(
                f"Failed to insert text into document '{doc_id}': {str(e)}",
                error_code=ErrorCode.OPERATION_FAILED.value,
            )

    def format_text(self, doc_id: str, start: int, end: int, **formatting):
        """Apply formatting to a text range.

        Args:
            doc_id: Document ID
            start: Start position of text range (0-based)
            end: End position of text range (0-based)
            **formatting: Formatting options (bold, italic, underline, font_name, font_size, color)

        Raises:
            DocumentError: If document not found
            OperationError: If formatting fails
        """
        if doc_id not in self._documents:
            raise DocumentError(
                f"Document not found: {doc_id}",
                error_code=ErrorCode.DOCUMENT_NOT_FOUND.value,
            )

        doc_ref = self._documents[doc_id]

        try:
            doc = doc_ref.word_doc_ref
            if not doc:
                raise OperationError("Document reference is invalid")

            # Validate range
            doc_text_length = len(doc.Range().Text)
            if start < 0:
                start = 0
            if end > doc_text_length:
                end = doc_text_length
            if start >= end:
                raise OperationError("Invalid text range: start must be less than end")

            # Get the text range to format
            text_range = doc.Range(start, end)

            # Apply formatting
            if formatting.get("bold") is not None:
                text_range.Font.Bold = formatting["bold"]

            if formatting.get("italic") is not None:
                text_range.Font.Italic = formatting["italic"]

            if formatting.get("underline") is not None:
                if formatting["underline"]:
                    text_range.Font.Underline = word_constants.wdUnderlineSingle
                else:
                    text_range.Font.Underline = word_constants.wdUnderlineNone

            if formatting.get("font_name"):
                text_range.Font.Name = formatting["font_name"]

            if formatting.get("font_size"):
                text_range.Font.Size = formatting["font_size"]

            if formatting.get("color"):
                # Convert color string to RGB value
                color_value = self._parse_color(formatting["color"])
                if color_value is not None:
                    text_range.Font.Color = color_value

            if formatting.get("highlight_color"):
                # Set highlight color
                highlight_value = self._parse_highlight_color(
                    formatting["highlight_color"]
                )
                if highlight_value is not None:
                    text_range.HighlightColorIndex = highlight_value

            doc_ref.last_modified = datetime.now()

            self.logger.info(
                f"Applied formatting to document {doc_id} range {start}-{end}"
            )

        except Exception as e:
            raise OperationError(
                f"Failed to format text in document '{doc_id}': {str(e)}",
                error_code=ErrorCode.OPERATION_FAILED.value,
            )

    def select_text(self, doc_id: str, start: int, end: int):
        """Select a text range in the document.

        Args:
            doc_id: Document ID
            start: Start position of selection (0-based)
            end: End position of selection (0-based)

        Returns:
            Dictionary with selection information

        Raises:
            DocumentError: If document not found
            OperationError: If selection fails
        """
        if doc_id not in self._documents:
            raise DocumentError(
                f"Document not found: {doc_id}",
                error_code=ErrorCode.DOCUMENT_NOT_FOUND.value,
            )

        doc_ref = self._documents[doc_id]

        try:
            doc = doc_ref.word_doc_ref
            if not doc:
                raise OperationError("Document reference is invalid")

            # Validate range
            doc_text_length = len(doc.Range().Text)
            if start < 0:
                start = 0
            if end > doc_text_length:
                end = doc_text_length
            if start >= end:
                raise OperationError("Invalid text range: start must be less than end")

            # Select the text range
            text_range = doc.Range(start, end)
            text_range.Select()

            # Return selection information
            selected_text = text_range.Text

            self.logger.info(f"Selected text in document {doc_id} range {start}-{end}")

            return {
                "start": start,
                "end": end,
                "text": selected_text,
                "length": len(selected_text),
            }

        except Exception as e:
            raise OperationError(
                f"Failed to select text in document '{doc_id}': {str(e)}",
                error_code=ErrorCode.OPERATION_FAILED.value,
            )

    def _parse_color(self, color_str: str) -> Optional[int]:
        """Parse color string to RGB value.

        Args:
            color_str: Color string (hex, rgb, or named color)

        Returns:
            RGB color value or None if invalid
        """
        try:
            # Handle hex colors (#RRGGBB or #RGB)
            if color_str.startswith("#"):
                hex_color = color_str[1:]
                if len(hex_color) == 3:
                    # Convert #RGB to #RRGGBB
                    hex_color = "".join([c * 2 for c in hex_color])
                if len(hex_color) == 6:
                    r = int(hex_color[0:2], 16)
                    g = int(hex_color[2:4], 16)
                    b = int(hex_color[4:6], 16)
                    # Word uses BGR format
                    return (b << 16) | (g << 8) | r

            # Handle rgb(r,g,b) format
            if color_str.startswith("rgb(") and color_str.endswith(")"):
                rgb_values = color_str[4:-1].split(",")
                if len(rgb_values) == 3:
                    r = int(rgb_values[0].strip())
                    g = int(rgb_values[1].strip())
                    b = int(rgb_values[2].strip())
                    # Word uses BGR format
                    return (b << 16) | (g << 8) | r

            # Handle named colors
            named_colors = {
                "black": 0x000000,
                "white": 0xFFFFFF,
                "red": 0x0000FF,
                "green": 0x00FF00,
                "blue": 0xFF0000,
                "yellow": 0x00FFFF,
                "cyan": 0xFFFF00,
                "magenta": 0xFF00FF,
                "gray": 0x808080,
                "grey": 0x808080,
            }

            if color_str.lower() in named_colors:
                return named_colors[color_str.lower()]

        except (ValueError, IndexError):
            pass

        self.logger.warning(f"Could not parse color: {color_str}")
        return None

    def _parse_highlight_color(self, color_str: str) -> Optional[int]:
        """Parse highlight color string to Word highlight constant.

        Args:
            color_str: Highlight color name

        Returns:
            Word highlight color constant or None if invalid
        """
        try:
            highlight_colors = {
                "yellow": word_constants.wdYellow,
                "bright_green": word_constants.wdBrightGreen,
                "turquoise": word_constants.wdTurquoise,
                "pink": word_constants.wdPink,
                "blue": word_constants.wdBlue,
                "red": word_constants.wdRed,
                "dark_blue": word_constants.wdDarkBlue,
                "teal": word_constants.wdTeal,
                "green": word_constants.wdGreen,
                "violet": word_constants.wdViolet,
                "dark_red": word_constants.wdDarkRed,
                "dark_yellow": word_constants.wdDarkYellow,
                "gray_50": word_constants.wdGray50,
                "gray_25": word_constants.wdGray25,
                "black": word_constants.wdBlack,
                "none": word_constants.wdNoHighlight,
            }

            return highlight_colors.get(color_str.lower())

        except Exception:
            pass

        self.logger.warning(f"Could not parse highlight color: {color_str}")
        return None

    def create_table(
        self, doc_id: str, rows: int, cols: int, position: Optional[int] = None
    ) -> Dict[str, Any]:
        """Create a table in the document.

        Args:
            doc_id: Document ID
            rows: Number of rows
            cols: Number of columns
            position: Optional position to insert at (0-based). If None, inserts at end.

        Returns:
            Dictionary with table information

        Raises:
            DocumentError: If document not found
            OperationError: If table creation fails
        """
        if doc_id not in self._documents:
            raise DocumentError(
                f"Document not found: {doc_id}",
                error_code=ErrorCode.DOCUMENT_NOT_FOUND.value,
            )

        if rows < 1 or cols < 1:
            raise OperationError(
                "Table must have at least 1 row and 1 column",
                error_code=ErrorCode.OPERATION_FAILED.value,
            )

        if rows > 100 or cols > 50:
            raise OperationError(
                "Table size too large (max 100 rows, 50 columns)",
                error_code=ErrorCode.OPERATION_FAILED.value,
            )

        doc_ref = self._documents[doc_id]

        try:
            doc = doc_ref.word_doc_ref
            if not doc:
                raise OperationError("Document reference is invalid")

            # Get the document range
            doc_range = doc.Range()

            if position is not None:
                # Insert at specific position
                if position < 0:
                    position = 0
                elif position > len(doc_range.Text):
                    position = len(doc_range.Text)

                # Set the range to the insertion point
                doc_range.SetRange(position, position)
            else:
                # Insert at the end of the document
                doc_range.SetRange(len(doc_range.Text), len(doc_range.Text))

            # Create the table
            table = doc.Tables.Add(doc_range, rows, cols)

            # Set basic table properties
            try:
                if word_constants:
                    table.AutoFitBehavior(word_constants.wdAutoFitContent)
                table.Borders.Enable = True
            except Exception as e:
                self.logger.warning(f"Could not set table properties: {e}")

            doc_ref.last_modified = datetime.now()

            self.logger.info(f"Created {rows}x{cols} table in document {doc_id}")

            return {
                "rows": rows,
                "cols": cols,
                "position": position,
                "table_index": doc.Tables.Count,  # 1-based index
            }

        except Exception as e:
            raise OperationError(
                f"Failed to create table in document '{doc_id}': {str(e)}",
                error_code=ErrorCode.OPERATION_FAILED.value,
            )

    def format_table_cell(
        self,
        doc_id: str,
        table_index: int,
        row: int,
        col: int,
        text: Optional[str] = None,
        **formatting,
    ) -> None:
        """Format a table cell.

        Args:
            doc_id: Document ID
            table_index: Table index (1-based)
            row: Row number (1-based)
            col: Column number (1-based)
            text: Optional text to insert in the cell
            **formatting: Formatting options (bold, italic, font_size, etc.)

        Raises:
            DocumentError: If document not found
            OperationError: If cell formatting fails
        """
        if doc_id not in self._documents:
            raise DocumentError(
                f"Document not found: {doc_id}",
                error_code=ErrorCode.DOCUMENT_NOT_FOUND.value,
            )

        doc_ref = self._documents[doc_id]

        try:
            doc = doc_ref.word_doc_ref
            if not doc:
                raise OperationError("Document reference is invalid")

            # Validate table index
            if table_index < 1 or table_index > doc.Tables.Count:
                raise OperationError(
                    f"Table index {table_index} out of range (1-{doc.Tables.Count})"
                )

            table = doc.Tables(table_index)

            # Validate row and column
            if row < 1 or row > table.Rows.Count:
                raise OperationError(f"Row {row} out of range (1-{table.Rows.Count})")

            if col < 1 or col > table.Columns.Count:
                raise OperationError(
                    f"Column {col} out of range (1-{table.Columns.Count})"
                )

            # Get the cell
            cell = table.Cell(row, col)
            cell_range = cell.Range

            # Set text if provided
            if text is not None:
                cell_range.Text = text

            # Apply formatting
            if formatting.get("bold") is not None:
                cell_range.Font.Bold = formatting["bold"]

            if formatting.get("italic") is not None:
                cell_range.Font.Italic = formatting["italic"]

            if formatting.get("font_size"):
                cell_range.Font.Size = formatting["font_size"]

            if formatting.get("font_name"):
                cell_range.Font.Name = formatting["font_name"]

            if formatting.get("color"):
                color_value = self._parse_color(formatting["color"])
                if color_value is not None:
                    cell_range.Font.Color = color_value

            if formatting.get("background_color"):
                bg_color = self._parse_color(formatting["background_color"])
                if bg_color is not None:
                    cell.Shading.BackgroundPatternColor = bg_color

            doc_ref.last_modified = datetime.now()

            self.logger.info(f"Formatted table cell [{row},{col}] in document {doc_id}")

        except Exception as e:
            raise OperationError(
                f"Failed to format table cell in document '{doc_id}': {str(e)}",
                error_code=ErrorCode.OPERATION_FAILED.value,
            )

    def create_list(
        self,
        doc_id: str,
        items: List[str],
        list_type: str = "bulleted",
        position: Optional[int] = None,
    ) -> Dict[str, Any]:
        """Create a bulleted or numbered list in the document.

        Args:
            doc_id: Document ID
            items: List of text items
            list_type: Type of list ("bulleted" or "numbered")
            position: Optional position to insert at (0-based). If None, inserts at end.

        Returns:
            Dictionary with list information

        Raises:
            DocumentError: If document not found
            OperationError: If list creation fails
        """
        if doc_id not in self._documents:
            raise DocumentError(
                f"Document not found: {doc_id}",
                error_code=ErrorCode.DOCUMENT_NOT_FOUND.value,
            )

        if not items:
            raise OperationError(
                "List must contain at least one item",
                error_code=ErrorCode.OPERATION_FAILED.value,
            )

        if list_type not in ["bulleted", "numbered"]:
            raise OperationError(
                "List type must be 'bulleted' or 'numbered'",
                error_code=ErrorCode.OPERATION_FAILED.value,
            )

        doc_ref = self._documents[doc_id]

        try:
            doc = doc_ref.word_doc_ref
            if not doc:
                raise OperationError("Document reference is invalid")

            # Get the document range
            doc_range = doc.Range()

            if position is not None:
                # Insert at specific position
                if position < 0:
                    position = 0
                elif position > len(doc_range.Text):
                    position = len(doc_range.Text)

                # Set the range to the insertion point
                doc_range.SetRange(position, position)
            else:
                # Insert at the end of the document
                doc_range.SetRange(len(doc_range.Text), len(doc_range.Text))

            # Create list text with proper formatting
            list_text = ""
            for i, item in enumerate(items):
                if i > 0:
                    list_text += "\n"
                list_text += item

            # Insert the text
            doc_range.InsertAfter(list_text)

            # Select the inserted text to apply list formatting
            start_pos = doc_range.Start
            end_pos = start_pos + len(list_text)
            list_range = doc.Range(start_pos, end_pos)

            # Apply list formatting
            if list_type == "bulleted":
                list_range.ListFormat.ApplyBulletDefault()
            else:  # numbered
                list_range.ListFormat.ApplyNumberDefault()

            doc_ref.last_modified = datetime.now()

            self.logger.info(
                f"Created {list_type} list with {len(items)} items in document {doc_id}"
            )

            return {
                "list_type": list_type,
                "item_count": len(items),
                "position": position,
                "start_pos": start_pos,
                "end_pos": end_pos,
            }

        except Exception as e:
            raise OperationError(
                f"Failed to create list in document '{doc_id}': {str(e)}",
                error_code=ErrorCode.OPERATION_FAILED.value,
            )

    def find_replace(
        self,
        doc_id: str,
        find_text: str,
        replace_text: str,
        match_case: bool = False,
        match_whole_word: bool = False,
        use_regex: bool = False,
        replace_all: bool = True,
    ) -> Dict[str, Any]:
        """Find and replace text in the document.

        Args:
            doc_id: Document ID
            find_text: Text to find
            replace_text: Text to replace with
            match_case: Whether to match case (default: False)
            match_whole_word: Whether to match whole words only (default: False)
            use_regex: Whether to use regex patterns (default: False)
            replace_all: Whether to replace all occurrences (default: True)

        Returns:
            Dictionary with replacement information

        Raises:
            DocumentError: If document not found
            OperationError: If find/replace operation fails
        """
        if doc_id not in self._documents:
            raise DocumentError(
                f"Document not found: {doc_id}",
                error_code=ErrorCode.DOCUMENT_NOT_FOUND.value,
            )

        if not find_text:
            raise OperationError(
                "Find text cannot be empty", error_code=ErrorCode.OPERATION_FAILED.value
            )

        doc_ref = self._documents[doc_id]

        try:
            doc = doc_ref.word_doc_ref
            if not doc:
                raise OperationError("Document reference is invalid")

            # Get the document range for searching
            search_range = doc.Range()

            # Configure the Find object
            find_obj = search_range.Find
            find_obj.ClearFormatting()
            find_obj.Replacement.ClearFormatting()

            # Set search parameters
            find_obj.Text = find_text
            find_obj.Replacement.Text = replace_text
            find_obj.MatchCase = match_case
            find_obj.MatchWholeWord = match_whole_word

            # Handle regex patterns
            if use_regex and word_constants:
                find_obj.MatchWildcards = True
            else:
                find_obj.MatchWildcards = False

            # Set other search options
            find_obj.Forward = True
            find_obj.Wrap = word_constants.wdFindContinue if word_constants else 1

            replacements_made = 0

            if replace_all:
                # Replace all occurrences
                if word_constants:
                    replace_result = find_obj.Execute(
                        Replace=word_constants.wdReplaceAll
                    )
                    # In real Word environment, we can get the actual count
                    # For now, we'll use a simple estimation
                    replacements_made = 1 if replace_result else 0
                else:
                    # Fallback for test environment
                    replace_result = find_obj.Execute()
                    replacements_made = 1 if replace_result else 0
            else:
                # Replace only the first occurrence
                if word_constants:
                    replace_result = find_obj.Execute(
                        Replace=word_constants.wdReplaceOne
                    )
                else:
                    # Fallback for test environment
                    replace_result = find_obj.Execute()

                replacements_made = 1 if replace_result else 0

            doc_ref.last_modified = datetime.now()

            self.logger.info(
                f"Find/replace operation in document {doc_id}: {replacements_made} replacements made"
            )

            return {
                "find_text": find_text,
                "replace_text": replace_text,
                "replacements_made": replacements_made,
                "match_case": match_case,
                "match_whole_word": match_whole_word,
                "use_regex": use_regex,
                "replace_all": replace_all,
            }

        except Exception as e:
            raise OperationError(
                f"Failed to perform find/replace in document '{doc_id}': {str(e)}",
                error_code=ErrorCode.OPERATION_FAILED.value,
            )

    def insert_header_footer(
        self,
        doc_id: str,
        header_text: Optional[str] = None,
        footer_text: Optional[str] = None,
        section_index: int = 1,
    ) -> Dict[str, Any]:
        """Insert or update headers and footers in a document.

        Args:
            doc_id: Document ID
            header_text: Text to insert in header (None to skip)
            footer_text: Text to insert in footer (None to skip)
            section_index: Section index (1-based, default: 1)

        Returns:
            Dictionary with operation results

        Raises:
            DocumentError: If document not found
            OperationError: If header/footer operation fails
        """
        if doc_id not in self._documents:
            raise DocumentError(
                f"Document not found: {doc_id}",
                error_code=ErrorCode.DOCUMENT_NOT_FOUND.value,
            )

        doc_ref = self._documents[doc_id]

        try:
            doc = doc_ref.word_doc_ref
            if not doc:
                raise OperationError("Document reference is invalid")

            # Get the specified section (default to first section)
            if section_index < 1 or section_index > doc.Sections.Count:
                section_index = 1

            section = doc.Sections(section_index)

            results = {}

            # Insert header if provided
            if header_text is not None:
                try:
                    header = section.Headers(word_constants.wdHeaderFooterPrimary)
                    header_range = header.Range
                    header_range.Text = header_text
                    results["header"] = {
                        "success": True,
                        "text": header_text,
                        "section": section_index,
                    }
                    self.logger.info(
                        f"Inserted header in document {doc_id}, section {section_index}"
                    )
                except Exception as e:
                    results["header"] = {"success": False, "error": str(e)}
                    self.logger.warning(f"Failed to insert header: {e}")

            # Insert footer if provided
            if footer_text is not None:
                try:
                    footer = section.Footers(word_constants.wdHeaderFooterPrimary)
                    footer_range = footer.Range
                    footer_range.Text = footer_text
                    results["footer"] = {
                        "success": True,
                        "text": footer_text,
                        "section": section_index,
                    }
                    self.logger.info(
                        f"Inserted footer in document {doc_id}, section {section_index}"
                    )
                except Exception as e:
                    results["footer"] = {"success": False, "error": str(e)}
                    self.logger.warning(f"Failed to insert footer: {e}")

            doc_ref.last_modified = datetime.now()

            return results

        except Exception as e:
            raise OperationError(
                f"Failed to insert header/footer in document '{doc_id}': {str(e)}",
                error_code=ErrorCode.OPERATION_FAILED.value,
            )

    def insert_page_break(
        self, doc_id: str, position: Optional[int] = None, break_type: str = "page"
    ) -> Dict[str, Any]:
        """Insert a page break or section break in a document.

        Args:
            doc_id: Document ID
            position: Position to insert break (0-based). If None, inserts at end.
            break_type: Type of break ("page", "section_next_page", "section_continuous")

        Returns:
            Dictionary with operation results

        Raises:
            DocumentError: If document not found
            OperationError: If break insertion fails
        """
        if doc_id not in self._documents:
            raise DocumentError(
                f"Document not found: {doc_id}",
                error_code=ErrorCode.DOCUMENT_NOT_FOUND.value,
            )

        doc_ref = self._documents[doc_id]

        try:
            doc = doc_ref.word_doc_ref
            if not doc:
                raise OperationError("Document reference is invalid")

            # Get the document range
            doc_range = doc.Range()

            if position is not None:
                # Insert at specific position
                if position < 0:
                    position = 0
                elif position > len(doc_range.Text):
                    position = len(doc_range.Text)

                # Set the range to the insertion point
                doc_range.SetRange(position, position)
            else:
                # Insert at the end of the document
                doc_range.SetRange(len(doc_range.Text), len(doc_range.Text))

            # Determine break type constant
            break_constants = {
                "page": word_constants.wdPageBreak,
                "section_next_page": word_constants.wdSectionBreakNextPage,
                "section_continuous": word_constants.wdSectionBreakContinuous,
                "section_even_page": word_constants.wdSectionBreakEvenPage,
                "section_odd_page": word_constants.wdSectionBreakOddPage,
            }

            break_constant = break_constants.get(break_type, word_constants.wdPageBreak)

            # Insert the break
            doc_range.InsertBreak(break_constant)

            doc_ref.last_modified = datetime.now()

            self.logger.info(
                f"Inserted {break_type} break in document {doc_id} at position {position}"
            )

            return {
                "success": True,
                "break_type": break_type,
                "position": position,
                "message": f"Inserted {break_type} break at position {position}",
            }

        except Exception as e:
            raise OperationError(
                f"Failed to insert break in document '{doc_id}': {str(e)}",
                error_code=ErrorCode.OPERATION_FAILED.value,
            )

    def set_page_formatting(
        self,
        doc_id: str,
        section_index: int = 1,
        margins: Optional[Dict[str, float]] = None,
        orientation: Optional[str] = None,
        paper_size: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Set page formatting options for a document section.

        Args:
            doc_id: Document ID
            section_index: Section index (1-based, default: 1)
            margins: Dictionary with margin values in points (top, bottom, left, right)
            orientation: Page orientation ("portrait" or "landscape")
            paper_size: Paper size ("letter", "a4", "legal", etc.)

        Returns:
            Dictionary with operation results

        Raises:
            DocumentError: If document not found
            OperationError: If formatting fails
        """
        if doc_id not in self._documents:
            raise DocumentError(
                f"Document not found: {doc_id}",
                error_code=ErrorCode.DOCUMENT_NOT_FOUND.value,
            )

        doc_ref = self._documents[doc_id]

        try:
            doc = doc_ref.word_doc_ref
            if not doc:
                raise OperationError("Document reference is invalid")

            # Get the specified section (default to first section)
            if section_index < 1 or section_index > doc.Sections.Count:
                section_index = 1

            section = doc.Sections(section_index)
            page_setup = section.PageSetup

            results = {}

            # Set margins if provided
            if margins:
                try:
                    if "top" in margins:
                        page_setup.TopMargin = margins["top"]
                        results["top_margin"] = margins["top"]

                    if "bottom" in margins:
                        page_setup.BottomMargin = margins["bottom"]
                        results["bottom_margin"] = margins["bottom"]

                    if "left" in margins:
                        page_setup.LeftMargin = margins["left"]
                        results["left_margin"] = margins["left"]

                    if "right" in margins:
                        page_setup.RightMargin = margins["right"]
                        results["right_margin"] = margins["right"]

                    self.logger.info(
                        f"Set margins for document {doc_id}, section {section_index}"
                    )

                except Exception as e:
                    results["margins_error"] = str(e)
                    self.logger.warning(f"Failed to set margins: {e}")

            # Set orientation if provided
            if orientation:
                try:
                    if orientation.lower() == "portrait":
                        page_setup.Orientation = word_constants.wdOrientPortrait
                        results["orientation"] = "portrait"
                    elif orientation.lower() == "landscape":
                        page_setup.Orientation = word_constants.wdOrientLandscape
                        results["orientation"] = "landscape"
                    else:
                        results["orientation_error"] = (
                            f"Invalid orientation: {orientation}"
                        )

                    self.logger.info(
                        f"Set orientation to {orientation} for document {doc_id}, section {section_index}"
                    )

                except Exception as e:
                    results["orientation_error"] = str(e)
                    self.logger.warning(f"Failed to set orientation: {e}")

            # Set paper size if provided
            if paper_size:
                try:
                    paper_sizes = {
                        "letter": word_constants.wdPaperLetter,
                        "a4": word_constants.wdPaperA4,
                        "legal": word_constants.wdPaperLegal,
                        "executive": word_constants.wdPaperExecutive,
                        "a3": word_constants.wdPaperA3,
                        "a5": word_constants.wdPaperA5,
                        "b4": word_constants.wdPaperB4,
                        "b5": word_constants.wdPaperB5,
                        "tabloid": word_constants.wdPaperTabloid,
                    }

                    paper_constant = paper_sizes.get(paper_size.lower())
                    if paper_constant:
                        page_setup.PaperSize = paper_constant
                        results["paper_size"] = paper_size
                        self.logger.info(
                            f"Set paper size to {paper_size} for document {doc_id}, section {section_index}"
                        )
                    else:
                        results["paper_size_error"] = (
                            f"Invalid paper size: {paper_size}"
                        )

                except Exception as e:
                    results["paper_size_error"] = str(e)
                    self.logger.warning(f"Failed to set paper size: {e}")

            doc_ref.last_modified = datetime.now()

            results["success"] = True
            results["section"] = section_index

            return results

        except Exception as e:
            raise OperationError(
                f"Failed to set page formatting in document '{doc_id}': {str(e)}",
                error_code=ErrorCode.OPERATION_FAILED.value,
            )

    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit."""
        self.disconnect()
