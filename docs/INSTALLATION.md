# Installation Guide

This guide covers the installation and setup of Word MCP Server for use with Claude and other LLMs.

## System Requirements

### Operating System
- **Windows 10/11** (required for Microsoft Word COM interface)
- Windows Server 2016 or later (for server deployments)

### Software Requirements
- **Python 3.8 or higher**
- **Microsoft Word 2016 or later** (Office 365, Office 2019, or Office 2021)
- **pip** (Python package installer)

### Hardware Requirements
- **RAM**: Minimum 4GB, recommended 8GB or more
- **Storage**: At least 100MB free space for the package
- **CPU**: Any modern x64 processor

## Installation Methods

### Method 1: Install from PyPI (Recommended)

```bash
# Install the latest stable version
pip install word-mcp-server

# Install a specific version
pip install word-mcp-server==0.1.0

# Install with development dependencies (for contributors)
pip install word-mcp-server[dev]
```

### Method 2: Install from Source

```bash
# Clone the repository
git clone https://github.com/word-mcp-server/word-mcp-server.git
cd word-mcp-server

# Install in development mode
pip install -e .

# Or install with development dependencies
pip install -e .[dev]
```

### Method 3: Install from Wheel

```bash
# Download the wheel file from GitHub releases
# Then install it
pip install word_mcp_server-0.1.0-py3-none-any.whl
```

## Post-Installation Setup

### 1. Verify Installation

```bash
# Check if the package is installed correctly
word-mcp-server --version

# Check system requirements
word-mcp-server --check-requirements
```

### 2. Create Configuration File

```bash
# Create a default configuration file
word-mcp-server --create-config

# Create config in a specific location
word-mcp-server --create-config -o /path/to/config.yaml
```

### 3. Run Setup Wizard (Optional)

```bash
# Run the interactive setup wizard
word-mcp-server --setup
```

The setup wizard will:
- Check system requirements
- Configure Word settings
- Set up security preferences
- Test the Word connection
- Create a sample configuration

## Configuration

### Basic Configuration

The default configuration file (`config.yaml`) contains:

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

### Advanced Configuration

For advanced users, you can customize:

- **Server settings**: Host, port, timeouts
- **Word behavior**: Visibility, auto-save, backup options
- **Security**: File access restrictions, macro settings
- **Logging**: Log levels, file rotation
- **Performance**: Concurrent document limits

See [Configuration Guide](CONFIGURATION.md) for detailed options.

## Claude Integration

### 1. Configure MCP in Claude

Add the following to your Claude MCP configuration:

```json
{
  "mcpServers": {
    "word-office": {
      "command": "word-mcp-server",
      "args": [],
      "env": {
        "WORD_MCP_CONFIG": "config.yaml"
      }
    }
  }
}
```

### 2. Start the Server

```bash
# Start with default configuration
word-mcp-server

# Start with custom configuration
word-mcp-server --config /path/to/config.yaml

# Start with verbose logging
word-mcp-server --verbose
```

### 3. Test the Connection

Once the server is running, you can test it in Claude:

```
Can you create a new Word document and add some text?
```

## Deployment Options

### Local Development

For local development and testing:

```bash
# Install in development mode
pip install -e .[dev]

# Run with verbose logging
word-mcp-server --verbose --config dev-config.yaml
```

### Production Deployment

#### Option 1: Windows Service

```bash
# Install as Windows service (requires admin privileges)
word-mcp-server --install-service

# Start the service
net start WordMCPServer

# Stop the service
net stop WordMCPServer

# Uninstall the service
word-mcp-server --uninstall-service
```

#### Option 2: Task Scheduler

1. Open Windows Task Scheduler
2. Create a new task
3. Set trigger to "At startup"
4. Set action to run: `word-mcp-server --config C:\path\to\config.yaml`
5. Configure to run as a service account

#### Option 3: Docker (Experimental)

```dockerfile
# Note: Requires Windows containers and Word installation
FROM mcr.microsoft.com/windows/servercore:ltsc2022

# Install Python and Word MCP Server
RUN pip install word-mcp-server

# Copy configuration
COPY config.yaml /app/config.yaml

# Expose port
EXPOSE 8080

# Start server
CMD ["word-mcp-server", "--config", "/app/config.yaml"]
```

## Troubleshooting

### Common Issues

#### 1. "Word not found" Error

```bash
# Check if Word is installed and registered
word-mcp-server --check-requirements

# Try launching Word manually
start winword
```

#### 2. COM Interface Errors

```bash
# Re-register Word COM components (run as administrator)
regsvr32 /i msword.olb

# Or repair Office installation
```

#### 3. Permission Errors

```bash
# Run as administrator (if needed)
# Or configure allowed_paths in config.yaml
```

#### 4. Port Already in Use

```yaml
# Change port in config.yaml
server:
  port: 8081  # Use different port
```

### Getting Help

1. **Check logs**: Look at `word_mcp_server.log` for detailed error messages
2. **Run diagnostics**: Use `word-mcp-server --check-requirements`
3. **Verbose mode**: Run with `--verbose` for detailed output
4. **GitHub Issues**: Report bugs at [GitHub Issues](https://github.com/word-mcp-server/word-mcp-server/issues)

## Uninstallation

### Remove Package

```bash
# Uninstall the package
pip uninstall word-mcp-server

# Remove configuration files
word-mcp-server --uninstall
```

### Clean Uninstall

```bash
# Remove all traces
pip uninstall word-mcp-server
rm -rf ~/.word-mcp-server/
rm config.yaml
rm word_mcp_server.log
```

## Next Steps

After successful installation:

1. **Read the [User Guide](USER_GUIDE.md)** for usage instructions
2. **Check [Examples](../examples/)** for sample code
3. **Review [API Reference](API_REFERENCE.md)** for detailed documentation
4. **Join the community** for support and updates

## Version Compatibility

| Word MCP Server | Python | Word | Claude |
|----------------|--------|------|--------|
| 0.1.x          | 3.8+   | 2016+ | All versions |

## Security Considerations

- **File Access**: Configure `allowed_paths` to restrict file access
- **Macros**: Keep `enable_macros: false` unless needed
- **Network**: Use `localhost` binding for local-only access
- **Updates**: Keep the package updated for security fixes

For detailed security guidelines, see [Security Guide](SECURITY.md).