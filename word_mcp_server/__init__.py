"""
Word Office MCP Server

A Model Context Protocol server that enables LLMs like Claude to interact with Microsoft Word.
"""

__version__ = "0.1.0"
__author__ = "Word MCP Server Team"
__description__ = "MCP server for Microsoft Word automation"

from .config import ConfigManager
from .server import WordMCPServer
from .word import DocumentManager, WordController

__all__ = ["WordMCPServer", "WordController", "DocumentManager", "ConfigManager"]
