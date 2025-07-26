# Word MCP Server Troubleshooting Guide

This guide helps you diagnose and resolve common issues when using the Word MCP Server.

## Table of Contents

1. [Installation Issues](#installation-issues)
2. [Connection Problems](#connection-problems)
3. [Document Operation Errors](#document-operation-errors)
4. [Performance Issues](#performance-issues)
5. [Configuration Problems](#configuration-problems)
6. [Claude Integration Issues](#claude-integration-issues)
7. [Debugging Tips](#debugging-tips)
8. [Getting Help](#getting-help)

## Installation Issues

### Problem: Package Installation Fails

**Symptoms:**
- `pip install word-mcp-server` fails
- Missing dependencies errors
- Python version compatibility issues

**Solutions:**

1. **Check Python Version:**
   ```bash
   python --version
   # Ensure Python 3.8 or higher
   ```

2. **Upgrade pip:**
   ```bash
   python -m pip install --upgrade pip
   ```

3. **Install with specific Python version:**
   ```bash
   python3.9 -m pip install word-mcp-server
   ```

4. **Install in virtual environment:**
   ```bash
   python -m venv word-mcp-env
   source word-mcp-env/bin/activate  # On Windows: word-mcp-env\Scripts\activate
   pip install word-mcp-server
   ```

### Problem: pywin32 Installation Issues (Windows)

**Symptoms:**
- `ImportError: No module named 'win32com'`
- COM interface not available errors

**Solutions:**

1. **Install pywin32 manually:**
   ```bash
   pip install pywin32
   ```

2. **Run post-install script:**
   ```bash
   python Scripts/pywin32_postinstall.py -install
   ```

3. **For conda environments:**
   ```bash
   conda install pywin32
   ```

### Problem: Microsoft Word Not Detected

**Symptoms:**
- "Microsoft Word not found" errors
- COM registration issues

**Solutions:**

1. **Verify Word installation:**
   - Open Word manually to ensure it works
   - Check Word version (2016 or later recommended)

2. **Re-register Word COM components:**
   ```cmd
   # Run as Administrator
   regsvr32 "C:\Program Files\Microsoft Office\root\Office16\MSWORD.OLB"
   ```

3. **Repair Office installation:**
   - Go to Control Panel > Programs
   - Find Microsoft Office, click "Change"
   - Select "Quick Repair" or "Online Repair"

## Connection Problems

### Problem: Cannot Connect to Word Application

**Error Code:** `WORD_CONNECTION_FAILED`

**Symptoms:**
- Server fails to start
- "Could not establish connection to Word application" errors

**Solutions:**

1. **Check Word is not running:**
   ```bash
   # Close all Word instances before starting server
   taskkill /f /im WINWORD.EXE  # Windows
   ```

2. **Run with administrator privileges:**
   - Right-click command prompt, "Run as administrator"
   - Start the MCP server from elevated prompt

3. **Check COM security settings:**
   ```bash
   # Run dcomcnfg.exe as administrator
   # Navigate to Component Services > Computers > My Computer > DCOM Config
   # Find "Microsoft Word 97 - 2003 Document"
   # Right-click > Properties > Security
   # Ensure appropriate permissions for your user
   ```

4. **Verify Word automation settings:**
   - Open Word > File > Options > Trust Center > Trust Center Settings
   - Macro Settings > Enable all macros (temporarily for testing)

### Problem: Connection Lost During Operation

**Error Code:** `WORD_CRASHED`

**Symptoms:**
- Operations fail mid-execution
- "Word application became unresponsive" errors

**Solutions:**

1. **Enable automatic recovery:**
   ```yaml
   # In config.yaml
   word:
     auto_launch: true
     recovery_enabled: true
     max_retry_attempts: 3
   ```

2. **Increase timeout values:**
   ```yaml
   server:
     timeout_seconds: 60
     operation_timeout: 30
   ```

3. **Check system resources:**
   - Ensure sufficient RAM (4GB+ recommended)
   - Close unnecessary applications
   - Check disk space

## Document Operation Errors

### Problem: Document Not Found

**Error Code:** `DOCUMENT_NOT_FOUND`

**Symptoms:**
- "Document not found" when opening files
- Path-related errors

**Solutions:**

1. **Verify file path:**
   ```python
   import os
   path = "/path/to/document.docx"
   print(f"File exists: {os.path.exists(path)}")
   print(f"Absolute path: {os.path.abspath(path)}")
   ```

2. **Check file permissions:**
   ```python
   import os
   path = "/path/to/document.docx"
   print(f"Readable: {os.access(path, os.R_OK)}")
   print(f"Writable: {os.access(path, os.W_OK)}")
   ```

3. **Use absolute paths:**
   ```python
   from pathlib import Path
   doc_path = Path("/path/to/document.docx").resolve()
   ```

### Problem: Access Denied Errors

**Error Code:** `DOCUMENT_ACCESS_DENIED`

**Symptoms:**
- Cannot open or save documents
- Permission-related errors

**Solutions:**

1. **Check file is not open elsewhere:**
   - Close document in Word
   - Check for background Word processes

2. **Verify file permissions:**
   - Right-click file > Properties > Security
   - Ensure your user has Full Control

3. **Run as administrator:**
   - Start command prompt as administrator
   - Run MCP server from elevated prompt

4. **Check antivirus software:**
   - Temporarily disable real-time protection
   - Add Word and Python to antivirus exceptions

### Problem: Invalid Document ID

**Error Code:** `INVALID_DOCUMENT_ID`

**Symptoms:**
- Operations fail with "Document not found" for valid-looking IDs
- Document references become invalid

**Solutions:**

1. **Check document lifecycle:**
   ```python
   # Ensure document wasn't closed
   doc_id = await server._handle_create_document()
   # ... perform operations ...
   # Don't use doc_id after closing
   await server._handle_close_document(doc_id=doc_id)
   ```

2. **Verify document ID format:**
   ```python
   # Document IDs should be UUIDs
   import uuid
   print(f"Valid UUID: {uuid.UUID(doc_id)}")
   ```

3. **List active documents:**
   ```python
   # Check what documents are currently open
   active_docs = server.word_controller.list_documents()
   print(f"Active documents: {list(active_docs.keys())}")
   ```

## Performance Issues

### Problem: Slow Document Operations

**Symptoms:**
- Operations take longer than expected
- Timeouts during large document processing

**Solutions:**

1. **Optimize document size:**
   - Break large documents into smaller sections
   - Process documents in chunks

2. **Adjust configuration:**
   ```yaml
   server:
     max_concurrent_docs: 5  # Reduce if system is slow
     timeout_seconds: 120    # Increase for large documents
   
   word:
     visible: false          # Keep Word hidden for better performance
     save_on_exit: false     # Disable auto-save for temp operations
   ```

3. **Monitor system resources:**
   ```python
   import psutil
   print(f"CPU usage: {psutil.cpu_percent()}%")
   print(f"Memory usage: {psutil.virtual_memory().percent}%")
   ```

### Problem: Memory Usage Issues

**Symptoms:**
- High memory consumption
- System becomes unresponsive
- Out of memory errors

**Solutions:**

1. **Close documents when finished:**
   ```python
   # Always close documents
   await server._handle_close_document(doc_id=doc_id, save=False)
   ```

2. **Limit concurrent operations:**
   ```yaml
   server:
     max_concurrent_docs: 3  # Reduce concurrent document limit
   ```

3. **Enable garbage collection:**
   ```python
   import gc
   gc.collect()  # Force garbage collection periodically
   ```

## Configuration Problems

### Problem: Configuration File Not Found

**Symptoms:**
- "Configuration file not found" errors
- Server uses default settings

**Solutions:**

1. **Create configuration file:**
   ```bash
   mkdir -p ~/.word-mcp-server
   cp config.yaml.example ~/.word-mcp-server/config.yaml
   ```

2. **Specify config path:**
   ```bash
   export WORD_MCP_CONFIG="/path/to/config.yaml"
   word-mcp-server
   ```

3. **Use default configuration:**
   ```python
   from word_mcp_server.config.config_manager import ConfigManager
   config = ConfigManager()  # Uses built-in defaults
   ```

### Problem: Invalid Configuration Values

**Symptoms:**
- "Invalid configuration" errors
- Server fails to start

**Solutions:**

1. **Validate configuration:**
   ```python
   from word_mcp_server.config.config_manager import ConfigManager
   try:
       config = ConfigManager("/path/to/config.yaml")
       print("Configuration is valid")
   except Exception as e:
       print(f"Configuration error: {e}")
   ```

2. **Check YAML syntax:**
   ```bash
   python -c "import yaml; yaml.safe_load(open('config.yaml'))"
   ```

3. **Use configuration template:**
   ```yaml
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
   ```

## Claude Integration Issues

### Problem: Claude Cannot Connect to MCP Server

**Symptoms:**
- Claude shows "MCP server not available"
- Connection timeout errors

**Solutions:**

1. **Verify MCP server configuration:**
   ```json
   {
     "mcpServers": {
       "word-office": {
         "command": "python",
         "args": ["-m", "word_mcp_server"],
         "env": {
           "WORD_MCP_CONFIG": "config.yaml"
         }
       }
     }
   }
   ```

2. **Check server is running:**
   ```bash
   # Test server manually
   python -m word_mcp_server --test
   ```

3. **Verify Python path:**
   ```bash
   which python
   python -c "import word_mcp_server; print('OK')"
   ```

### Problem: Tools Not Available in Claude

**Symptoms:**
- Claude doesn't see Word tools
- "Tool not found" errors

**Solutions:**

1. **Restart Claude:**
   - Close and reopen Claude application
   - Clear Claude's cache if possible

2. **Check tool registration:**
   ```python
   from word_mcp_server.server.mcp_server import WordMCPServer
   server = WordMCPServer(config_manager)
   print(f"Available tools: {list(server.tools.keys())}")
   ```

3. **Verify MCP protocol version:**
   - Ensure compatible MCP library versions
   - Update both Claude and word-mcp-server

## Debugging Tips

### Enable Debug Logging

```yaml
logging:
  level: "DEBUG"
  file: "debug.log"
  console: true
```

### Test Individual Components

```python
# Test Word connection
from word_mcp_server.word.controller import WordController
from word_mcp_server.config.models import WordConfig

config = WordConfig()
controller = WordController(config)
print(f"Connected: {controller.connect_to_word()}")

# Test document operations
doc_id = controller.create_document()
print(f"Document created: {doc_id}")
```

### Monitor System Resources

```python
import psutil
import time

def monitor_resources():
    while True:
        cpu = psutil.cpu_percent()
        memory = psutil.virtual_memory().percent
        print(f"CPU: {cpu}%, Memory: {memory}%")
        time.sleep(5)
```

### Check COM Objects

```python
import win32com.client

try:
    word_app = win32com.client.GetActiveObject("Word.Application")
    print(f"Word version: {word_app.Version}")
    print(f"Documents open: {word_app.Documents.Count}")
except:
    print("No active Word application found")
```

## Common Error Messages and Solutions

| Error Message | Cause | Solution |
|---------------|-------|----------|
| "COM interface not available" | pywin32 not installed or registered | Install pywin32, run post-install script |
| "Word application not found" | Word not installed or not in PATH | Install Word, check installation |
| "Access denied to document" | File permissions or file in use | Check permissions, close file in Word |
| "Document not found" | Invalid file path | Verify path exists and is accessible |
| "Connection timeout" | Word not responding | Restart Word, check system resources |
| "Invalid document ID" | Document closed or invalid ID | Check document lifecycle, verify ID |
| "Operation failed" | General Word operation error | Check Word status, retry operation |

## Getting Help

### Log Files

Check these log files for detailed error information:
- `word_mcp_server.log` - Main server log
- `debug.log` - Debug information (if enabled)
- Windows Event Viewer - System-level errors

### Diagnostic Information

When reporting issues, include:

```python
import sys
import platform
import word_mcp_server

print(f"Python version: {sys.version}")
print(f"Platform: {platform.platform()}")
print(f"Word MCP Server version: {word_mcp_server.__version__}")
print(f"Working directory: {os.getcwd()}")
```

### Support Channels

1. **GitHub Issues**: Report bugs and feature requests
2. **Documentation**: Check the latest documentation
3. **Community Forums**: Ask questions and share solutions
4. **Stack Overflow**: Tag questions with `word-mcp-server`

### Before Reporting Issues

1. Check this troubleshooting guide
2. Search existing GitHub issues
3. Test with minimal configuration
4. Gather diagnostic information
5. Create a minimal reproduction case

Remember to never include sensitive information (passwords, personal documents) in bug reports or logs.