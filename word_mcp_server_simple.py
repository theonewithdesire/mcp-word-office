#!/usr/bin/env python3
"""
Simplified Word MCP Server for testing with Claude/Kiro.
This version focuses on basic MCP protocol compatibility.
"""

import asyncio
import json
import sys
from typing import Any, Dict, List

# Simple MCP server implementation
class SimpleMCPServer:
    def __init__(self):
        self.tools = self._create_word_tools()
    
    def _create_word_tools(self):
        """Create the list of Word tools."""
        return [
            {
                "name": "create_document",
                "description": "Create a new Word document",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "title": {
                            "type": "string",
                            "description": "Optional title for the document"
                        }
                    }
                }
            },
            {
                "name": "insert_text",
                "description": "Insert text into a Word document",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "doc_id": {"type": "string", "description": "Document ID"},
                        "text": {"type": "string", "description": "Text to insert"}
                    },
                    "required": ["doc_id", "text"]
                }
            },
            {
                "name": "format_text",
                "description": "Apply formatting to text in a Word document",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "doc_id": {"type": "string", "description": "Document ID"},
                        "start": {"type": "integer", "description": "Start position"},
                        "end": {"type": "integer", "description": "End position"},
                        "bold": {"type": "boolean", "description": "Apply bold formatting"}
                    },
                    "required": ["doc_id", "start", "end"]
                }
            },
            {
                "name": "save_document",
                "description": "Save a Word document",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "doc_id": {"type": "string", "description": "Document ID"},
                        "path": {"type": "string", "description": "Optional save path"}
                    },
                    "required": ["doc_id"]
                }
            },
            {
                "name": "read_document",
                "description": "Read content from a Word document",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "path": {"type": "string", "description": "Path to document"}
                    },
                    "required": ["path"]
                }
            },
            {
                "name": "create_table",
                "description": "Create a table in a Word document",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "doc_id": {"type": "string", "description": "Document ID"},
                        "rows": {"type": "integer", "description": "Number of rows"},
                        "cols": {"type": "integer", "description": "Number of columns"}
                    },
                    "required": ["doc_id", "rows", "cols"]
                }
            },
            {
                "name": "create_list",
                "description": "Create a bulleted or numbered list",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "doc_id": {"type": "string", "description": "Document ID"},
                        "items": {"type": "array", "items": {"type": "string"}},
                        "list_type": {"type": "string", "enum": ["bulleted", "numbered"]}
                    },
                    "required": ["doc_id", "items"]
                }
            },
            {
                "name": "find_replace",
                "description": "Find and replace text in a Word document",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "doc_id": {"type": "string", "description": "Document ID"},
                        "find_text": {"type": "string", "description": "Text to find"},
                        "replace_text": {"type": "string", "description": "Replacement text"}
                    },
                    "required": ["doc_id", "find_text", "replace_text"]
                }
            }
        ]
    
    async def handle_request(self, request: Dict[str, Any]) -> Dict[str, Any]:
        """Handle MCP requests."""
        method = request.get("method")
        request_id = request.get("id")
        params = request.get("params", {})
        
        try:
            if method == "initialize":
                return {
                    "jsonrpc": "2.0",
                    "id": request_id,
                    "result": {
                        "protocolVersion": "2024-11-05",
                        "capabilities": {
                            "tools": {}
                        },
                        "serverInfo": {
                            "name": "word-mcp-server",
                            "version": "0.1.0"
                        }
                    }
                }
            
            elif method == "tools/list":
                return {
                    "jsonrpc": "2.0",
                    "id": request_id,
                    "result": {
                        "tools": self.tools
                    }
                }
            
            elif method == "tools/call":
                tool_name = params.get("name")
                arguments = params.get("arguments", {})
                
                # Simulate tool execution (will fail on Mac, but shows connection works)
                if tool_name in [tool["name"] for tool in self.tools]:
                    return {
                        "jsonrpc": "2.0",
                        "id": request_id,
                        "result": {
                            "content": [
                                {
                                    "type": "text",
                                    "text": f"❌ Tool '{tool_name}' failed: Microsoft Word COM automation requires Windows. However, the MCP connection is working correctly! Arguments received: {arguments}"
                                }
                            ]
                        }
                    }
                else:
                    return {
                        "jsonrpc": "2.0",
                        "id": request_id,
                        "error": {
                            "code": -32601,
                            "message": f"Unknown tool: {tool_name}"
                        }
                    }
            
            else:
                return {
                    "jsonrpc": "2.0",
                    "id": request_id,
                    "error": {
                        "code": -32601,
                        "message": f"Method not found: {method}"
                    }
                }
        
        except Exception as e:
            return {
                "jsonrpc": "2.0",
                "id": request_id,
                "error": {
                    "code": -32603,
                    "message": f"Internal error: {str(e)}"
                }
            }
    
    async def run(self):
        """Run the MCP server with stdio transport."""
        print("🚀 Simple Word MCP Server starting...", file=sys.stderr)
        print(f"📋 Registered {len(self.tools)} Word tools", file=sys.stderr)
        
        try:
            while True:
                # Read JSON-RPC request from stdin
                line = await asyncio.get_event_loop().run_in_executor(
                    None, sys.stdin.readline
                )
                
                if not line:
                    break
                
                try:
                    request = json.loads(line.strip())
                    response = await self.handle_request(request)
                    
                    # Send response to stdout
                    print(json.dumps(response), flush=True)
                    
                except json.JSONDecodeError as e:
                    print(f"❌ JSON decode error: {e}", file=sys.stderr)
                    continue
                    
        except KeyboardInterrupt:
            print("🛑 Server stopped by user", file=sys.stderr)
        except Exception as e:
            print(f"❌ Server error: {e}", file=sys.stderr)

async def main():
    """Main entry point."""
    server = SimpleMCPServer()
    await server.run()

if __name__ == "__main__":
    asyncio.run(main())