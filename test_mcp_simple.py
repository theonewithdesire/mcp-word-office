#!/usr/bin/env python3
"""
Simple test to check if MCP server starts and shows any errors.
"""

import subprocess
import sys
import time

def test_server_startup():
    """Test if the server starts without errors."""
    print("Testing MCP server startup...")
    
    try:
        # Start the server and capture both stdout and stderr
        process = subprocess.Popen(
            ["word-mcp-server"],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )
        
        # Give it a moment to start
        time.sleep(2)
        
        # Check if process is still running
        if process.poll() is None:
            print("✅ Server started successfully and is running")
            
            # Try to send a simple message
            try:
                process.stdin.write('{"jsonrpc": "2.0", "id": 1, "method": "initialize", "params": {}}\n')
                process.stdin.flush()
                
                # Wait a bit for response
                time.sleep(1)
                
                # Check for any output
                process.stdout.settimeout(1)  # This won't work, but let's try another approach
                
            except Exception as e:
                print(f"Error sending message: {e}")
            
            # Terminate the process
            process.terminate()
            stdout, stderr = process.communicate(timeout=5)
            
            print(f"STDOUT: {stdout}")
            print(f"STDERR: {stderr}")
            
        else:
            # Process exited, get the output
            stdout, stderr = process.communicate()
            print(f"❌ Server exited with code: {process.returncode}")
            print(f"STDOUT: {stdout}")
            print(f"STDERR: {stderr}")
            return False
            
    except Exception as e:
        print(f"❌ Error starting server: {e}")
        return False
    
    return True

if __name__ == "__main__":
    success = test_server_startup()
    sys.exit(0 if success else 1)