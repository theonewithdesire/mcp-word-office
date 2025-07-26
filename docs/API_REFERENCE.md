# Word MCP Server API Reference

This document provides comprehensive documentation for all tools available in the Word MCP Server.

## Table of Contents

1. [Document Lifecycle Operations](#document-lifecycle-operations)
2. [Text Operations](#text-operations)
3. [Formatting Operations](#formatting-operations)
4. [Table Operations](#table-operations)
5. [List Operations](#list-operations)
6. [Find and Replace Operations](#find-and-replace-operations)
7. [Header and Footer Operations](#header-and-footer-operations)
8. [Page Formatting Operations](#page-formatting-operations)
9. [Document Reading Operations](#document-reading-operations)
10. [Error Handling](#error-handling)

## Document Lifecycle Operations

### create_document

Creates a new Word document.

**Parameters:**
- `title` (string, optional): Title for the document

**Returns:**
- Document ID (string): Unique identifier for the created document

**Example:**
```python
doc_id = await server._handle_create_document(title="My Document")
```

### open_document

Opens an existing Word document.

**Parameters:**
- `path` (string, required): Path to the document file

**Returns:**
- Document ID (string): Unique identifier for the opened document

**Example:**
```python
doc_id = await server._handle_open_document(path="/path/to/document.docx")
```

### save_document

Saves a Word document.

**Parameters:**
- `doc_id` (string, required): Document ID to save
- `path` (string, optional): Path to save to (if not provided, saves to current location)

**Returns:**
- None

**Example:**
```python
await server._handle_save_document(doc_id="doc-123", path="/path/to/save.docx")
```

### close_document

Closes a Word document.

**Parameters:**
- `doc_id` (string, required): Document ID to close
- `save` (boolean, optional, default: True): Whether to save before closing

**Returns:**
- None

**Example:**
```python
await server._handle_close_document(doc_id="doc-123", save=True)
```

## Text Operations

### insert_text

Inserts text at a specified position in the document.

**Parameters:**
- `doc_id` (string, required): Document ID
- `text` (string, required): Text to insert
- `position` (integer, optional): Position to insert text (0 = beginning)

**Returns:**
- Dictionary with operation results:
  - `success` (boolean): Whether operation succeeded
  - `doc_id` (string): Document ID
  - `text` (string): Inserted text
  - `position` (integer): Position where text was inserted
  - `length` (integer): Length of inserted text

**Example:**
```python
result = await server._handle_insert_text(
    doc_id="doc-123",
    text="Hello, World!",
    position=0
)
```

### select_text

Selects a range of text in the document.

**Parameters:**
- `doc_id` (string, required): Document ID
- `start` (integer, required): Start position of selection
- `end` (integer, required): End position of selection

**Returns:**
- Dictionary with selection results

**Example:**
```python
result = await server._handle_select_text(
    doc_id="doc-123",
    start=0,
    end=10
)
```

## Formatting Operations

### format_text

Applies formatting to a range of text.

**Parameters:**
- `doc_id` (string, required): Document ID
- `start` (integer, required): Start position of text to format
- `end` (integer, required): End position of text to format
- `bold` (boolean, optional): Apply bold formatting
- `italic` (boolean, optional): Apply italic formatting
- `underline` (boolean, optional): Apply underline formatting
- `font_name` (string, optional): Font name (e.g., "Arial", "Times New Roman")
- `font_size` (integer, optional): Font size in points
- `color` (string, optional): Text color (e.g., "red", "blue", "#FF0000")
- `highlight_color` (string, optional): Highlight color

**Returns:**
- Dictionary with formatting results:
  - `success` (boolean): Whether operation succeeded
  - `start` (integer): Start position of formatted text
  - `end` (integer): End position of formatted text
  - `formatting` (object): Applied formatting options

**Example:**
```python
result = await server._handle_format_text(
    doc_id="doc-123",
    start=0,
    end=13,
    bold=True,
    font_size=16,
    color="blue"
)
```

## Table Operations

### create_table

Creates a table in the document.

**Parameters:**
- `doc_id` (string, required): Document ID
- `rows` (integer, required): Number of rows
- `cols` (integer, required): Number of columns
- `position` (integer, optional): Position to insert table

**Returns:**
- Dictionary with table creation results:
  - `success` (boolean): Whether operation succeeded
  - `rows` (integer): Number of rows created
  - `columns` (integer): Number of columns created
  - `table_id` (string): Unique identifier for the table

**Example:**
```python
result = await server._handle_create_table(
    doc_id="doc-123",
    rows=3,
    cols=4,
    position=100
)
```

### format_table_cell

Formats a specific table cell.

**Parameters:**
- `doc_id` (string, required): Document ID
- `table_index` (integer, required): Index of the table (0-based)
- `row` (integer, required): Row index (0-based)
- `col` (integer, required): Column index (0-based)
- Additional formatting parameters (same as format_text)

**Returns:**
- Dictionary with cell formatting results

**Example:**
```python
result = await server._handle_format_table_cell(
    doc_id="doc-123",
    table_index=0,
    row=0,
    col=0,
    bold=True,
    color="blue"
)
```

## List Operations

### create_list

Creates a bulleted or numbered list.

**Parameters:**
- `doc_id` (string, required): Document ID
- `items` (array of strings, required): List items
- `list_type` (string, optional, default: "bulleted"): Type of list ("bulleted" or "numbered")
- `position` (integer, optional): Position to insert list

**Returns:**
- Dictionary with list creation results:
  - `success` (boolean): Whether operation succeeded
  - `items` (array): Created list items
  - `list_type` (string): Type of list created

**Example:**
```python
result = await server._handle_create_list(
    doc_id="doc-123",
    items=["Item 1", "Item 2", "Item 3"],
    list_type="bulleted",
    position=50
)
```

## Find and Replace Operations

### find_replace

Performs find and replace operations in the document.

**Parameters:**
- `doc_id` (string, required): Document ID
- `find_text` (string, required): Text to find
- `replace_text` (string, required): Text to replace with
- `match_case` (boolean, optional, default: False): Whether to match case
- `whole_words` (boolean, optional, default: False): Whether to match whole words only
- `use_regex` (boolean, optional, default: False): Whether to use regular expressions

**Returns:**
- Dictionary with replace results:
  - `success` (boolean): Whether operation succeeded
  - `replacements` (integer): Number of replacements made
  - `find_text` (string): Text that was searched for
  - `replace_text` (string): Replacement text

**Example:**
```python
result = await server._handle_find_replace(
    doc_id="doc-123",
    find_text="old text",
    replace_text="new text",
    match_case=False
)
```

## Header and Footer Operations

### insert_header_footer

Inserts or updates headers and footers.

**Parameters:**
- `doc_id` (string, required): Document ID
- `header_text` (string, optional): Header text
- `footer_text` (string, optional): Footer text
- `section_index` (integer, optional, default: 1): Section index (1-based)

**Returns:**
- Dictionary with header/footer results:
  - `header` (object, optional): Header operation results
  - `footer` (object, optional): Footer operation results

**Example:**
```python
result = await server._handle_insert_header_footer(
    doc_id="doc-123",
    header_text="Document Title",
    footer_text="Page ",
    section_index=1
)
```

## Page Formatting Operations

### insert_page_break

Inserts a page break or section break.

**Parameters:**
- `doc_id` (string, required): Document ID
- `position` (integer, optional): Position to insert break (None for end of document)
- `break_type` (string, optional, default: "page"): Type of break ("page", "section_next_page", "section_continuous")

**Returns:**
- Dictionary with break insertion results:
  - `success` (boolean): Whether operation succeeded
  - `break_type` (string): Type of break inserted
  - `position` (integer): Position where break was inserted

**Example:**
```python
result = await server._handle_insert_page_break(
    doc_id="doc-123",
    position=100,
    break_type="page"
)
```

### set_page_formatting

Sets page formatting options.

**Parameters:**
- `doc_id` (string, required): Document ID
- `section_index` (integer, optional, default: 1): Section index (1-based)
- `margins` (object, optional): Margin settings
  - `top` (number): Top margin in points
  - `bottom` (number): Bottom margin in points
  - `left` (number): Left margin in points
  - `right` (number): Right margin in points
- `orientation` (string, optional): Page orientation ("portrait" or "landscape")
- `paper_size` (string, optional): Paper size ("letter", "a4", "legal", etc.)

**Returns:**
- Dictionary with formatting results

**Example:**
```python
result = await server._handle_set_page_formatting(
    doc_id="doc-123",
    section_index=1,
    margins={"top": 72, "bottom": 72, "left": 90, "right": 90},
    orientation="portrait",
    paper_size="letter"
)
```

## Document Reading Operations

### read_document

Reads and extracts content from a Word document.

**Parameters:**
- `path` (string, required): Path to the document file

**Returns:**
- Dictionary with document content:
  - `text` (string): Extracted text content
  - `paragraphs` (array): Array of paragraph texts
  - `structure` (object): Document structure information

**Example:**
```python
result = await server._handle_read_document(path="/path/to/document.docx")
```

### get_document_info

Retrieves metadata and information about a document.

**Parameters:**
- `path` (string, required): Path to the document file

**Returns:**
- Dictionary with document information:
  - `metadata` (object): Document metadata (title, author, created date, etc.)
  - `properties` (object): Document properties

**Example:**
```python
result = await server._handle_get_document_info(path="/path/to/document.docx")
```

### get_document_statistics

Calculates statistics for a document.

**Parameters:**
- `path` (string, required): Path to the document file

**Returns:**
- Dictionary with document statistics:
  - `word_count` (integer): Number of words
  - `character_count` (integer): Number of characters
  - `paragraph_count` (integer): Number of paragraphs
  - `page_count` (integer): Number of pages

**Example:**
```python
result = await server._handle_get_document_statistics(path="/path/to/document.docx")
```

### extract_comments

Extracts comments from a document.

**Parameters:**
- `path` (string, required): Path to the document file

**Returns:**
- Dictionary with comments:
  - `comments` (array): Array of comment objects

**Example:**
```python
result = await server._handle_extract_comments(path="/path/to/document.docx")
```

## Error Handling

The Word MCP Server provides comprehensive error handling with specific error codes and helpful messages.

### Error Response Format

All operations return error information in a consistent format:

```json
{
  "success": false,
  "error": "Error message describing what went wrong",
  "error_code": "SPECIFIC_ERROR_CODE",
  "details": "Additional details about the error",
  "suggestions": ["Suggestion 1", "Suggestion 2"]
}
```

### Common Error Codes

- `WORD_CONNECTION_FAILED`: Could not connect to Word application
- `DOCUMENT_NOT_FOUND`: Specified document file does not exist
- `DOCUMENT_ACCESS_DENIED`: Permission denied accessing document
- `INVALID_DOCUMENT_ID`: Document ID is not valid or document is closed
- `OPERATION_FAILED`: General operation failure
- `INVALID_PARAMETERS`: Invalid or missing required parameters

### Error Recovery

The server includes automatic error recovery mechanisms:

1. **Connection Recovery**: Automatically attempts to reconnect to Word if connection is lost
2. **Graceful Degradation**: Falls back to alternative methods when primary operations fail
3. **Retry Logic**: Automatically retries failed operations with exponential backoff

### Best Practices

1. **Always check operation results**: Verify the `success` field before proceeding
2. **Handle errors gracefully**: Provide meaningful feedback to users
3. **Use appropriate timeouts**: Don't wait indefinitely for operations to complete
4. **Clean up resources**: Always close documents when finished
5. **Validate inputs**: Check parameters before sending to the server

## Usage Examples

### Complete Document Workflow

```python
async def create_formatted_document():
    # Create document
    doc_id = await server._handle_create_document(title="My Report")
    
    # Add title
    await server._handle_insert_text(
        doc_id=doc_id,
        text="Annual Report 2024",
        position=0
    )
    
    # Format title
    await server._handle_format_text(
        doc_id=doc_id,
        start=0,
        end=18,
        bold=True,
        font_size=16
    )
    
    # Add content
    await server._handle_insert_text(
        doc_id=doc_id,
        text="\n\nThis is the content of the report...",
        position=18
    )
    
    # Create table
    await server._handle_create_table(
        doc_id=doc_id,
        rows=3,
        cols=2,
        position=50
    )
    
    # Save and close
    await server._handle_save_document(
        doc_id=doc_id,
        path="/path/to/report.docx"
    )
    await server._handle_close_document(doc_id=doc_id)
```

### Document Analysis Workflow

```python
async def analyze_document():
    # Read document content
    content = await server._handle_read_document(
        path="/path/to/document.docx"
    )
    
    # Get metadata
    info = await server._handle_get_document_info(
        path="/path/to/document.docx"
    )
    
    # Get statistics
    stats = await server._handle_get_document_statistics(
        path="/path/to/document.docx"
    )
    
    # Extract comments
    comments = await server._handle_extract_comments(
        path="/path/to/document.docx"
    )
    
    return {
        "content": content,
        "info": info,
        "stats": stats,
        "comments": comments
    }
```

This API reference provides comprehensive documentation for all available operations in the Word MCP Server. For more examples and usage patterns, see the examples directory in the project repository.