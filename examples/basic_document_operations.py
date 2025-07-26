#!/usr/bin/env python3
"""
Basic Document Operations Example

This example demonstrates basic document operations using the Word MCP server:
- Creating a new document
- Adding text content
- Formatting text
- Saving the document

Usage:
    python examples/basic_document_operations.py
"""

import asyncio
import sys
import os
from pathlib import Path

# Add the project root to the Python path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from word_mcp_server.server.mcp_server import WordMCPServer
from word_mcp_server.config.config_manager import ConfigManager


async def basic_document_example():
    """Demonstrate basic document operations."""
    print("=== Word MCP Server - Basic Document Operations Example ===\n")
    
    try:
        # Initialize the MCP server
        print("1. Initializing Word MCP Server...")
        config_manager = ConfigManager()
        server = WordMCPServer(config_manager)
        print("   ✓ Server initialized successfully\n")
        
        # Create a new document
        print("2. Creating a new document...")
        doc_id = await server._handle_create_document(title="Example Document")
        print(f"   ✓ Document created with ID: {doc_id}\n")
        
        # Add a title
        print("3. Adding document title...")
        title_result = await server._handle_insert_text(
            doc_id=doc_id,
            text="Word MCP Server Example Document",
            position=0
        )
        print(f"   ✓ Title added: {title_result['text']}\n")
        
        # Format the title as bold
        print("4. Formatting title as bold...")
        format_result = await server._handle_format_text(
            doc_id=doc_id,
            start=0,
            end=len("Word MCP Server Example Document"),
            bold=True,
            font_size=16
        )
        print("   ✓ Title formatted as bold\n")
        
        # Add a paragraph
        print("5. Adding content paragraph...")
        paragraph_text = "\n\nThis document was created using the Word MCP Server, which enables LLMs like Claude to interact with Microsoft Word through the Model Context Protocol."
        paragraph_result = await server._handle_insert_text(
            doc_id=doc_id,
            text=paragraph_text,
            position=len("Word MCP Server Example Document")
        )
        print(f"   ✓ Paragraph added: {len(paragraph_text)} characters\n")
        
        # Add a bulleted list
        print("6. Adding a bulleted list...")
        list_items = [
            "Create and edit Word documents",
            "Format text with various styles",
            "Insert tables and lists",
            "Manage headers and footers",
            "Perform find and replace operations"
        ]
        
        list_result = await server._handle_create_list(
            doc_id=doc_id,
            items=list_items,
            list_type="bulleted",
            position=len("Word MCP Server Example Document") + len(paragraph_text)
        )
        print(f"   ✓ Bulleted list added with {len(list_items)} items\n")
        
        # Save the document
        print("7. Saving the document...")
        save_path = project_root / "examples" / "output" / "basic_example.docx"
        save_path.parent.mkdir(exist_ok=True)
        
        await server._handle_save_document(
            doc_id=doc_id,
            path=str(save_path)
        )
        print(f"   ✓ Document saved to: {save_path}\n")
        
        # Close the document
        print("8. Closing the document...")
        await server._handle_close_document(doc_id=doc_id, save=True)
        print("   ✓ Document closed\n")
        
        print("=== Example completed successfully! ===")
        print(f"Check the output file at: {save_path}")
        
    except Exception as e:
        print(f"❌ Error occurred: {e}")
        print(f"Error type: {type(e).__name__}")
        import traceback
        traceback.print_exc()


async def demonstrate_error_handling():
    """Demonstrate error handling in the MCP server."""
    print("\n=== Error Handling Demonstration ===\n")
    
    try:
        config_manager = ConfigManager()
        server = WordMCPServer(config_manager)
        
        # Try to operate on a non-existent document
        print("1. Testing error handling with invalid document ID...")
        try:
            await server._handle_insert_text(
                doc_id="invalid-doc-id",
                text="This should fail",
                position=0
            )
        except Exception as e:
            print(f"   ✓ Expected error caught: {e}\n")
        
        # Try to open a non-existent file
        print("2. Testing error handling with non-existent file...")
        try:
            await server._handle_open_document(path="/nonexistent/file.docx")
        except Exception as e:
            print(f"   ✓ Expected error caught: {e}\n")
        
        print("=== Error handling demonstration completed ===")
        
    except Exception as e:
        print(f"❌ Unexpected error: {e}")


if __name__ == "__main__":
    print("Starting Word MCP Server Basic Example...")
    print("Make sure Microsoft Word is installed and accessible.\n")
    
    # Run the basic example
    asyncio.run(basic_document_example())
    
    # Demonstrate error handling
    asyncio.run(demonstrate_error_handling())