# Word Office MCP Server

A Model Context Protocol (MCP) server that enables LLMs like Claude to interact with Microsoft Word applications through Python automation.

## Features

- **Word Automation**: Control Microsoft Word through COM interface
- **Document Operations**: Create, open, save, and manipulate Word documents
- **Text Processing**: Insert, format, and modify text content
- **Advanced Features**: Tables, lists, headers/footers, find/replace
- **Document Reading**: Extract content and metadata from existing documents
- **MCP Protocol**: Full compatibility with Claude and other MCP clients

## Requirements

- Python 3.8 or higher
- Microsoft Word 2016 or later (Windows only)
- Windows operating system (for COM interface)

## Installation

```bash
pip install word-mcp-server
```

Or install from source:

```bash
git clone <repository-url>
cd word-mcp-server
pip install -e .
```

## Quick Start

1. Create a configuration file:
```bash
word-mcp-server --create-config
```

2. Start the server:
```bash
word-mcp-server
```

3. Configure Claude to use the MCP server by adding to your MCP configuration:
```json
{
  "mcpServers": {
    "word-office": {
      "command": "word-mcp-server",
      "args": []
    }
  }
}
```

## Configuration

The server uses a YAML configuration file (`config.yaml`) with the following structure:

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

security:
  allowed_paths:
    - "~/Documents"
    - "~/Desktop"
  enable_macros: false
  max_file_size_mb: 50
```

## Project Structure

```
word_mcp_server/
├── __init__.py              # Package initialization
├── main.py                  # Main entry point
├── config/                  # Configuration management
│   ├── __init__.py
│   ├── config_manager.py    # Configuration loader
│   └── models.py           # Configuration data models
├── server/                  # MCP server implementation
│   ├── __init__.py
│   └── mcp_server.py       # Main MCP server class
├── word/                    # Word application interface
│   ├── __init__.py
│   ├── controller.py       # COM interface controller
│   └── document_manager.py # Document file operations
└── utils/                   # Utility functions
    ├── __init__.py
    ├── errors.py           # Custom exceptions
    ├── logging.py          # Logging utilities
    └── validators.py       # Input validation
```

## Development Status

This project is currently under development. The following tasks have been completed:

- [x] Task 1: Project structure and core dependencies
- [ ] Task 2: MCP server foundation
- [ ] Task 3: Word COM interface controller
- [ ] Task 4: Basic document operations
- [ ] Task 5: Advanced document manipulation
- [ ] Task 6: Document reading capabilities
- [ ] Task 7: Headers, footers, and page formatting
- [ ] Task 8: Error handling and logging
- [ ] Task 9: Configuration and deployment
- [ ] Task 10: Advanced Word features
- [ ] Task 11: Testing and documentation
- [ ] Task 12: Packaging and distribution

## License

MIT License - see LICENSE file for details.

## Contributing

Contributions are welcome! Please read the contributing guidelines before submitting pull requests.

## Support

For issues and questions, please use the GitHub issue tracker.