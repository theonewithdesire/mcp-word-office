#!/usr/bin/env python3
"""
Test the Mac Word MCP server functionality.
"""

import asyncio
import subprocess
import json

async def test_mac_word_server():
    """Test the Mac Word MCP server."""
    print("🧪 Testing Mac Word MCP Server...")
    
    # Start the server
    process = await asyncio.create_subprocess_exec(
        "python3", "word_mcp_server_mac.py",
        stdin=asyncio.subprocess.PIPE,
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.PIPE
    )
    
    try:
        # Initialize
        init_msg = {"jsonrpc": "2.0", "id": 1, "method": "initialize", "params": {"protocolVersion": "2024-11-05", "capabilities": {}, "clientInfo": {"name": "test", "version": "1.0"}}}
        process.stdin.write((json.dumps(init_msg) + "\n").encode())
        await process.stdin.drain()
        
        init_response = await process.stdout.readline()
        print("✅ Initialized successfully")
        
        # Create document
        create_msg = {"jsonrpc": "2.0", "id": 2, "method": "tools/call", "params": {"name": "create_document", "arguments": {"title": "Test Document"}}}
        process.stdin.write((json.dumps(create_msg) + "\n").encode())
        await process.stdin.drain()
        
        create_response = await process.stdout.readline()
        create_data = json.loads(create_response.decode())
        print(f"📄 Create document: {create_data['result']['content'][0]['text']}")
        
        # Extract doc_id from response
        doc_id = create_data['result']['content'][0]['text'].split(': ')[1].split(' ')[0]
        
        # Add text
        text_msg = {"jsonrpc": "2.0", "id": 3, "method": "tools/call", "params": {"name": "insert_text", "arguments": {"doc_id": doc_id, "text": "Hello from Mac! This Word document was created using python-docx."}}}
        process.stdin.write((json.dumps(text_msg) + "\n").encode())
        await process.stdin.drain()
        
        text_response = await process.stdout.readline()
        text_data = json.loads(text_response.decode())
        print(f"📝 Insert text: {text_data['result']['content'][0]['text']}")
        
        # Save document
        save_msg = {"jsonrpc": "2.0", "id": 4, "method": "tools/call", "params": {"name": "save_document", "arguments": {"doc_id": doc_id, "filename": "test_mac_document"}}}
        process.stdin.write((json.dumps(save_msg) + "\n").encode())
        await process.stdin.drain()
        
        save_response = await process.stdout.readline()
        save_data = json.loads(save_response.decode())
        print(f"💾 Save document: {save_data['result']['content'][0]['text']}")
        
        return True
        
    except Exception as e:
        print(f"❌ Error: {e}")
        return False
        
    finally:
        process.terminate()
        await process.wait()

if __name__ == "__main__":
    success = asyncio.run(test_mac_word_server())
    if success:
        print("\n🎉 SUCCESS! Mac Word MCP Server is working perfectly!")
        print("📂 Check ~/Documents/MCP_Word_Documents/ for the created file")
    else:
        print("\n❌ Test failed")