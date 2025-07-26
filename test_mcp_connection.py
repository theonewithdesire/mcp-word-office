#!/usr/bin/env python3
"""
Test script to verify MCP server stdio connection.
"""

import asyncio
import subprocess
import json
import sys

async def test_mcp_server():
    """Test MCP server connection via stdio."""
    print("Testing MCP server stdio connection...")
    
    # Start the MCP server process
    process = await asyncio.create_subprocess_exec(
        "word-mcp-server",
        stdin=asyncio.subprocess.PIPE,
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.PIPE
    )
    
    try:
        # Send MCP initialization message
        init_message = {
            "jsonrpc": "2.0",
            "id": 1,
            "method": "initialize",
            "params": {
                "protocolVersion": "2024-11-05",
                "capabilities": {},
                "clientInfo": {
                    "name": "test-client",
                    "version": "1.0.0"
                }
            }
        }
        
        message_json = json.dumps(init_message) + "\n"
        print(f"Sending: {message_json.strip()}")
        
        # Send the message
        process.stdin.write(message_json.encode())
        await process.stdin.drain()
        
        # Wait for response (with timeout)
        try:
            stdout_data = await asyncio.wait_for(
                process.stdout.readline(), 
                timeout=10.0
            )
            
            if stdout_data:
                response = stdout_data.decode().strip()
                print(f"Received: {response}")
                
                # Try to parse as JSON
                try:
                    response_data = json.loads(response)
                    if "result" in response_data:
                        print("✅ MCP server responded successfully!")
                        print(f"Server capabilities: {response_data.get('result', {}).get('capabilities', {})}")
                        return True
                    else:
                        print(f"❌ Unexpected response format: {response_data}")
                        return False
                except json.JSONDecodeError as e:
                    print(f"❌ Invalid JSON response: {e}")
                    print(f"Raw response: {response}")
                    return False
            else:
                print("❌ No response received")
                return False
                
        except asyncio.TimeoutError:
            print("❌ Timeout waiting for response")
            return False
            
    except Exception as e:
        print(f"❌ Error during test: {e}")
        return False
        
    finally:
        # Clean up process
        try:
            process.terminate()
            await asyncio.wait_for(process.wait(), timeout=5.0)
        except:
            process.kill()
            await process.wait()

async def main():
    """Main test function."""
    print("Word MCP Server Connection Test")
    print("=" * 40)
    
    success = await test_mcp_server()
    
    if success:
        print("\n🎉 Test passed! MCP server is working correctly.")
        print("\nYou can now use this configuration in Claude/Kiro:")
        print("""
{
  "mcpServers": {
    "word-office": {
      "command": "word-mcp-server",
      "args": ["--config", "config.yaml"]
    }
  }
}
        """)
    else:
        print("\n❌ Test failed. Check the error messages above.")
        
        # Show stderr if available
        print("\nChecking for error messages...")
        try:
            process = await asyncio.create_subprocess_exec(
                "word-mcp-server", "--check-requirements",
                stdout=asyncio.subprocess.PIPE,
                stderr=asyncio.subprocess.PIPE
            )
            stdout, stderr = await process.communicate()
            
            if stdout:
                print(f"Requirements check output:\n{stdout.decode()}")
            if stderr:
                print(f"Error output:\n{stderr.decode()}")
                
        except Exception as e:
            print(f"Could not run requirements check: {e}")
    
    return success

if __name__ == "__main__":
    success = asyncio.run(main())
    sys.exit(0 if success else 1)