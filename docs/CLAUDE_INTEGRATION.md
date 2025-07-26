# Claude Integration Setup Guide

This guide walks you through setting up the Word MCP Server to work with Claude, enabling natural language control of Microsoft Word documents.

## Table of Contents

1. [Prerequisites](#prerequisites)
2. [Installation](#installation)
3. [Configuration](#configuration)
4. [Claude Setup](#claude-setup)
5. [Testing the Integration](#testing-the-integration)
6. [Usage Examples](#usage-examples)
7. [Troubleshooting](#troubleshooting)

## Prerequisites

### System Requirements

- **Operating System**: Windows 10/11 (required for COM interface)
- **Microsoft Word**: 2016 or later
- **Python**: 3.8 or higher
- **Claude**: Desktop application or API access

### Required Permissions

- Administrator access (for initial setup)
- Word automation permissions
- File system read/write access

## Installation

### Step 1: Install Word MCP Server

```bash
# Install from PyPI
pip install word-mcp-server

# Or install from source
git clone https://github.com/word-mcp-server/word-mcp-server.git
cd word-mcp-server
pip install -e .
```

### Step 2: Verify Installation

```bash
# Test the installation
python -m word_mcp_server --version

# Test Word connection
python -m word_mcp_server --test-connection
```

### Step 3: Install Dependencies

```bash
# Windows-specific dependencies
pip install pywin32

# Run post-install script (Windows)
python Scripts/pywin32_postinstall.py -install
```

## Configuration

### Step 1: Create Configuration File

Create a configuration file at `~/.word-mcp-server/config.yaml`:

```yaml
# Word MCP Server Configuration
server:
  host: "localhost"
  port: 8080
  max_concurrent_docs: 10
  timeout_seconds: 30

word:
  auto_launch: true
  visible: false
  save_on_exit: true
  backup_enabled: true

logging:
  level: "INFO"
  file: "word_mcp_server.log"
  max_size_mb: 100
  console: false

security:
  allowed_paths: 
    - "~/Documents"
    - "~/Desktop"
    - "C:/Users/%USERNAME%/Documents"
  enable_macros: false
  max_file_size_mb: 50
```

### Step 2: Test Configuration

```bash
# Validate configuration
python -c "from word_mcp_server.config.config_manager import ConfigManager; ConfigManager()"

# Test server startup
python -m word_mcp_server --config-test
```

## Claude Setup

### Step 1: Configure MCP Server in Claude

Add the Word MCP Server to Claude's configuration. The exact method depends on your Claude setup:

#### For Claude Desktop Application

1. Open Claude settings
2. Navigate to "MCP Servers" or "Extensions"
3. Add a new server with these settings:

```json
{
  "name": "word-office",
  "command": "python",
  "args": ["-m", "word_mcp_server"],
  "env": {
    "WORD_MCP_CONFIG": "~/.word-mcp-server/config.yaml"
  }
}
```

#### For Claude API Integration

Create an MCP configuration file (`mcp.json`):

```json
{
  "mcpServers": {
    "word-office": {
      "command": "python",
      "args": ["-m", "word_mcp_server"],
      "env": {
        "WORD_MCP_CONFIG": "config.yaml",
        "PYTHONPATH": "."
      },
      "cwd": "/path/to/word-mcp-server"
    }
  }
}
```

### Step 2: Verify Claude Configuration

1. Restart Claude
2. Check that Word tools are available
3. Look for tools like:
   - `create_document`
   - `insert_text`
   - `format_text`
   - `save_document`

## Testing the Integration

### Step 1: Basic Connection Test

Ask Claude to perform a simple operation:

```
"Create a new Word document with the title 'Test Document' and add some sample text."
```

Expected behavior:
- Claude should use the `create_document` tool
- Add text using `insert_text` tool
- Confirm successful creation

### Step 2: Advanced Operations Test

Try more complex operations:

```
"Create a document with a formatted title, add a table with 3 rows and 2 columns, and save it as 'advanced_test.docx'."
```

Expected behavior:
- Document creation
- Text formatting
- Table creation
- Document saving

### Step 3: Document Reading Test

Test document analysis capabilities:

```
"Read the document at C:/Users/[username]/Documents/sample.docx and summarize its contents."
```

Expected behavior:
- Document reading
- Content extraction
- Summary generation

## Usage Examples

### Creating a Formatted Report

```
"Create a professional report document with the following structure:
1. Title: 'Quarterly Sales Report Q4 2024' (bold, size 16)
2. Executive Summary section
3. A table showing sales data by region (4 rows, 3 columns)
4. Recommendations section with bulleted list
5. Add headers and footers
6. Save as 'quarterly_report.docx'"
```

### Document Analysis and Editing

```
"Open the document 'draft_report.docx', find all instances of 'Q3 2024' and replace them with 'Third Quarter 2024', then add a conclusion paragraph at the end."
```

### Batch Document Processing

```
"Create three similar documents for different departments (Sales, Marketing, HR) with customized content for each, including department-specific tables and formatting."
```

### Template-Based Document Creation

```
"Create a meeting minutes template with:
- Header with meeting title and date
- Attendees section (bulleted list)
- Agenda items (numbered list)
- Action items table
- Footer with page numbers
Save as 'meeting_minutes_template.docx'"
```

## Advanced Configuration

### Custom Tool Permissions

Restrict which tools Claude can use:

```yaml
security:
  allowed_tools:
    - "create_document"
    - "insert_text"
    - "format_text"
    - "save_document"
  restricted_tools:
    - "delete_document"
    - "run_macro"
```

### Performance Optimization

For better performance with Claude:

```yaml
server:
  max_concurrent_docs: 5  # Reduce for stability
  timeout_seconds: 60     # Increase for complex operations

word:
  visible: false          # Keep Word hidden
  save_on_exit: false     # Disable auto-save for temp docs

logging:
  level: "WARNING"        # Reduce log verbosity
  console: false          # Disable console output
```

### Network Configuration

If running Claude remotely:

```yaml
server:
  host: "0.0.0.0"        # Allow external connections
  port: 8080
  ssl_enabled: true       # Enable SSL for security
  ssl_cert: "cert.pem"
  ssl_key: "key.pem"
```

## Troubleshooting

### Common Issues

#### Claude Can't Find Word Tools

**Problem**: Claude doesn't show Word-related tools

**Solutions**:
1. Restart Claude application
2. Check MCP server configuration
3. Verify server is running:
   ```bash
   python -m word_mcp_server --status
   ```

#### Word Connection Failures

**Problem**: "Cannot connect to Word application"

**Solutions**:
1. Ensure Word is installed and licensed
2. Run as administrator
3. Check COM registration:
   ```cmd
   regsvr32 "C:\Program Files\Microsoft Office\root\Office16\MSWORD.OLB"
   ```

#### Permission Denied Errors

**Problem**: Cannot create or modify documents

**Solutions**:
1. Check file permissions
2. Verify allowed_paths in configuration
3. Run with elevated privileges

#### Slow Performance

**Problem**: Operations take too long

**Solutions**:
1. Reduce max_concurrent_docs
2. Increase timeout_seconds
3. Keep Word hidden (visible: false)
4. Close unnecessary applications

### Debug Mode

Enable debug logging for troubleshooting:

```yaml
logging:
  level: "DEBUG"
  file: "debug.log"
  console: true
```

Then check the debug log for detailed information about operations.

### Testing Individual Components

Test the MCP server independently:

```python
import asyncio
from word_mcp_server.server.mcp_server import WordMCPServer
from word_mcp_server.config.config_manager import ConfigManager

async def test_server():
    config = ConfigManager()
    server = WordMCPServer(config)
    
    # Test document creation
    doc_id = await server._handle_create_document(title="Test")
    print(f"Created document: {doc_id}")
    
    # Test text insertion
    result = await server._handle_insert_text(
        doc_id=doc_id,
        text="Hello, World!",
        position=0
    )
    print(f"Text inserted: {result}")

asyncio.run(test_server())
```

## Best Practices

### For Claude Interactions

1. **Be specific**: Provide clear instructions about formatting, structure, and content
2. **Use examples**: Show Claude what you want with concrete examples
3. **Break down complex tasks**: Split large operations into smaller steps
4. **Verify results**: Ask Claude to confirm operations completed successfully

### For Document Management

1. **Use descriptive filenames**: Help Claude understand document purposes
2. **Organize files**: Keep documents in well-structured directories
3. **Regular backups**: Enable backup_enabled in configuration
4. **Clean up**: Close documents when finished to free resources

### For Performance

1. **Batch operations**: Group related operations together
2. **Minimize formatting**: Apply formatting in batches rather than character-by-character
3. **Use templates**: Create reusable document templates
4. **Monitor resources**: Keep an eye on memory and CPU usage

## Security Considerations

### File Access Control

```yaml
security:
  allowed_paths:
    - "~/Documents"
    - "~/Desktop"
  blocked_extensions:
    - ".exe"
    - ".bat"
    - ".cmd"
```

### Macro Security

```yaml
word:
  enable_macros: false    # Disable macro execution
  trust_center_settings:
    disable_all_macros: true
    notify_for_signed_macros: false
```

### Network Security

- Use SSL/TLS for remote connections
- Implement authentication if needed
- Restrict network access to trusted sources

## Getting Help

If you encounter issues:

1. Check the [Troubleshooting Guide](TROUBLESHOOTING.md)
2. Review the [API Reference](API_REFERENCE.md)
3. Search existing GitHub issues
4. Create a new issue with:
   - System information
   - Configuration files (remove sensitive data)
   - Error messages and logs
   - Steps to reproduce

## Updates and Maintenance

### Keeping Up to Date

```bash
# Update Word MCP Server
pip install --upgrade word-mcp-server

# Check for updates
python -m word_mcp_server --check-updates
```

### Configuration Migration

When updating, check for configuration changes:

```bash
# Backup current configuration
cp config.yaml config.yaml.backup

# Generate new configuration template
python -m word_mcp_server --generate-config > config.yaml.new

# Merge configurations manually
```

This completes the Claude integration setup. With this configuration, you should be able to use natural language with Claude to control Microsoft Word documents through the MCP server.