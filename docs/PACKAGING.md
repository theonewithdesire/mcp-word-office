# Packaging and Distribution Guide

This document describes the packaging and distribution setup for Word MCP Server.

## Package Structure

The Word MCP Server is packaged as a standard Python package with the following structure:

```
word-mcp-server/
├── word_mcp_server/           # Main package
│   ├── __init__.py           # Package initialization
│   ├── __main__.py           # Module execution entry point
│   ├── main.py               # Main application logic
│   ├── config/               # Configuration management
│   ├── server/               # MCP server implementation
│   ├── utils/                # Utility modules
│   └── word/                 # Word automation modules
├── tests/                    # Test suite
├── docs/                     # Documentation
├── examples/                 # Example scripts
├── scripts/                  # Build and deployment scripts
├── pyproject.toml           # Package configuration
├── setup.py                 # Backward compatibility
├── requirements.txt         # Dependencies
├── MANIFEST.in              # Package manifest
├── LICENSE                  # MIT license
├── README.md                # Project documentation
└── CHANGELOG.md             # Version history
```

## Build System

The package uses modern Python packaging standards:

- **Build backend**: `setuptools.build_meta`
- **Configuration**: `pyproject.toml` (primary) + `setup.py` (compatibility)
- **Dependencies**: Declared in `pyproject.toml`
- **Entry points**: CLI command `word-mcp-server`

### Key Configuration

```toml
[project]
name = "word-mcp-server"
version = "0.1.0"
description = "A Model Context Protocol server for Microsoft Word automation"
license = "MIT"
requires-python = ">=3.8"

[project.scripts]
word-mcp-server = "word_mcp_server.main:main"
```

## Dependencies

### Core Dependencies
- `mcp>=0.1.0` - Model Context Protocol implementation
- `pywin32>=306` - Windows COM interface (Windows only)
- `python-docx>=1.1.0` - Word document manipulation
- `pydantic>=2.0.0` - Data validation and settings
- `PyYAML>=6.0` - Configuration file parsing
- `typing-extensions>=4.0.0` - Type hints compatibility
- `asyncio-mqtt>=0.13.0` - Async MQTT support
- `aiofiles>=23.0.0` - Async file operations

### Development Dependencies
- `pytest>=7.0.0` - Testing framework
- `pytest-asyncio>=0.21.0` - Async testing support
- `pytest-mock>=3.10.0` - Mocking utilities
- `black>=23.0.0` - Code formatting
- `isort>=5.12.0` - Import sorting
- `flake8>=6.0.0` - Linting
- `mypy>=1.0.0` - Type checking
- `coverage>=7.0.0` - Code coverage

## Build Process

### Manual Build

```bash
# Clean previous builds
rm -rf dist/ build/ *.egg-info/

# Build source distribution and wheel
python -m build

# Verify package
twine check dist/*
```

### Automated Build

Use the provided build script:

```bash
python scripts/build.py
```

This script performs:
1. Clean build artifacts
2. Run tests
3. Run linting and type checking
4. Build packages
5. Verify packages

## Distribution Formats

The package is distributed in two formats:

### 1. Source Distribution (sdist)
- **File**: `word_mcp_server-0.1.0.tar.gz`
- **Contains**: Source code, tests, documentation
- **Use case**: Development, custom builds

### 2. Wheel Distribution (bdist_wheel)
- **File**: `word_mcp_server-0.1.0-py3-none-any.whl`
- **Contains**: Compiled package, ready to install
- **Use case**: Production installation

## Installation Methods

### From PyPI (Recommended)
```bash
pip install word-mcp-server
```

### From Wheel File
```bash
pip install word_mcp_server-0.1.0-py3-none-any.whl
```

### From Source
```bash
git clone https://github.com/word-mcp-server/word-mcp-server.git
cd word-mcp-server
pip install -e .
```

## Entry Points

The package provides multiple ways to execute:

### 1. CLI Command
```bash
word-mcp-server --help
```

### 2. Module Execution
```bash
python -m word_mcp_server --help
```

### 3. Direct Import
```python
from word_mcp_server import WordMCPServer
```

## Testing

### Installation Testing

Test package installation in clean environments:

```bash
# Test local package
python scripts/test_installation.py --local

# Test built wheel
python scripts/test_installation.py --package dist/word_mcp_server-0.1.0-py3-none-any.whl

# Test multiple Python versions
python scripts/test_environments.py
```

### Test Results

The installation tests verify:
- ✅ Package installation via pip
- ✅ Entry point functionality
- ✅ Module execution
- ✅ Dependency resolution
- ✅ Configuration creation
- ⚠️ Requirements check (expected to fail on non-Windows)

## Release Process

### Automated Release

Use the release script:

```bash
python scripts/release.py
```

This handles:
1. Version updates
2. Changelog updates
3. Package building
4. PyPI upload
5. Git tagging

### Manual Release Steps

1. **Update version** in `pyproject.toml` and `main.py`
2. **Update CHANGELOG.md** with release notes
3. **Build package**: `python -m build`
4. **Test package**: `python scripts/test_installation.py`
5. **Upload to PyPI**: `twine upload dist/*`
6. **Create git tag**: `git tag v0.1.0 && git push origin v0.1.0`

## CI/CD Pipeline

### GitHub Actions

The project includes automated CI/CD workflows:

#### `.github/workflows/ci.yml`
- **Triggers**: Push to main/develop, pull requests
- **Jobs**: Test, build, security checks
- **Matrix**: Multiple Python versions and OS
- **Artifacts**: Built packages, test reports

#### `.github/workflows/release.yml`
- **Triggers**: Manual workflow dispatch
- **Jobs**: Version update, build, PyPI upload
- **Features**: Automated changelog, git tagging

### Workflow Features

- **Multi-platform testing**: Windows (primary), Ubuntu (limited)
- **Python version matrix**: 3.8, 3.9, 3.10, 3.11, 3.12
- **Code quality checks**: Black, isort, flake8, mypy
- **Security scanning**: Safety, Bandit
- **Coverage reporting**: Codecov integration
- **Automated publishing**: PyPI on release

## Package Metadata

### PyPI Information
- **Name**: `word-mcp-server`
- **License**: MIT
- **Platform**: Windows (primary), cross-platform (limited)
- **Python**: 3.8+
- **Keywords**: mcp, word, office, automation, claude, llm

### Classifiers
```python
classifiers = [
    "Development Status :: 3 - Alpha",
    "Intended Audience :: Developers",
    "Operating System :: Microsoft :: Windows",
    "Programming Language :: Python :: 3",
    "Topic :: Office/Business :: Office Suites",
    "Topic :: Software Development :: Libraries :: Python Modules",
]
```

## Quality Assurance

### Code Quality Tools
- **Black**: Code formatting (88 char line length)
- **isort**: Import sorting
- **flake8**: Linting and style checking
- **mypy**: Static type checking

### Testing Strategy
- **Unit tests**: Individual component testing
- **Integration tests**: End-to-end functionality
- **Installation tests**: Package installation verification
- **Environment tests**: Multiple Python version testing

### Security Measures
- **Dependency scanning**: Safety for known vulnerabilities
- **Code analysis**: Bandit for security issues
- **License compliance**: MIT license compatibility
- **Supply chain**: Verified dependencies only

## Troubleshooting

### Common Build Issues

1. **License classifier warnings**
   - **Solution**: Use `license = "MIT"` instead of `{text = "MIT"}`

2. **Missing files in package**
   - **Solution**: Update `MANIFEST.in` to include required files

3. **Entry point not working**
   - **Solution**: Verify `[project.scripts]` in `pyproject.toml`

4. **Dependency conflicts**
   - **Solution**: Use compatible version ranges in dependencies

### Installation Issues

1. **Windows-only dependencies**
   - **Expected**: `pywin32` only installs on Windows
   - **Solution**: Use conditional dependencies

2. **Python version compatibility**
   - **Check**: Ensure Python 3.8+ is used
   - **Solution**: Update Python or use compatible version

3. **Permission errors**
   - **Solution**: Use virtual environment or `--user` flag

## Best Practices

### Package Development
1. **Use semantic versioning** (MAJOR.MINOR.PATCH)
2. **Maintain CHANGELOG.md** for all releases
3. **Test in clean environments** before release
4. **Use type hints** for better code quality
5. **Document all public APIs** thoroughly

### Distribution
1. **Always test packages** before publishing
2. **Use wheel format** for faster installation
3. **Include comprehensive metadata** in pyproject.toml
4. **Provide clear installation instructions**
5. **Maintain backward compatibility** when possible

### Security
1. **Scan dependencies** for vulnerabilities
2. **Use minimal permissions** in CI/CD
3. **Validate all inputs** in package code
4. **Keep dependencies updated** regularly
5. **Follow security best practices** for automation

This packaging setup ensures reliable, secure, and maintainable distribution of the Word MCP Server package.