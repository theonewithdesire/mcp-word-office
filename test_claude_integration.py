#!/usr/bin/env python3
"""
Test script to verify MCP server works with proper initialization.
"""

import asyncio
import subprocess
import json
import sys

async def test_mcp_tools():
    """Test MCP server tool listing."""
    print("🧪 Testing MCP server tool listing...")
    
    # Start the MCP server process
    process = await asyncio.create_subprocess_exec(
        "word-mcp-server",
        stdin=asyncio.subprocess.PIPE,
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.PIPE
    )
    
    try:
        # Send proper MCP initialization
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
        
        # Send initialization
        message_json = json.dumps(init_message) + "\n"
        process.stdin.write(message_json.encode())
        await process.stdin.drain()
        
        # Wait for initialization response
        init_response = await asyncio.wait_for(process.stdout.readline(), timeout=5.0)
        print(f"📨 Init response: {init_response.decode().strip()}")
        
        # Send tools/list request
        tools_message = {
            "jsonrpc": "2.0",
            "id": 2,
            "method": "tools/list",
            "params": {}
        }
        
        message_json = json.dumps(tools_message) + "\n"
        process.stdin.write(message_json.encode())
        await process.stdin.drain()
        
        # Wait for tools response
        tools_response = await asyncio.wait_for(process.stdout.readline(), timeout=5.0)
        tools_data = json.loads(tools_response.decode().strip())
        
        if "result" in tools_data and "tools" in tools_data["result"]:
            tools = tools_data["result"]["tools"]
            print(f"✅ Found {len(tools)} tools:")
            for tool in tools[:5]:  # Show first 5 tools
                print(f"   📋 {tool['name']}: {tool['description']}")
            if len(tools) > 5:
                print(f"   ... and {len(tools) - 5} more tools")
            return True
        else:
            print(f"❌ Unexpected tools response: {tools_data}")
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
    print("🔧 Word MCP Server Integration Test")
    print("=" * 50)
    
    success = await test_mcp_tools()
    
    if success:
        print("\n🎉 SUCCESS! The MCP server is working correctly.")
        print("\n📝 What this means:")
        print("   ✅ Server starts and responds to MCP protocol")
        print("   ✅ Tools are properly registered and available")
        print("   ✅ Claude/Kiro should be able to see and use these tools")
        print("\n⚠️  Note: On Mac, Word operations will fail (Windows required)")
        print("   But Claude will still see the tools and can attempt to use them")
        
        print("\n🚀 Next steps:")
        print("   1. The MCP configuration is already updated")
        print("   2. Restart Kiro to reload MCP servers")
        print("   3. Ask Claude: 'What Word tools are available?'")
        print("   4. Try: 'Create a new Word document' (will show error but prove connection)")
        
    else:
        print("\n❌ Test failed. Check the error messages above.")
    
    return success

if __name__ == "__main__":
    success = asyncio.run(main())
    sys.exit(0 if success else 1)