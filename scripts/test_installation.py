#!/usr/bin/env python3
"""
Test script for Word MCP Server package installation in clean environments.
"""

import os
import sys
import subprocess
import tempfile
import shutil
import venv
from pathlib import Path


def run_command(cmd, cwd=None, capture_output=True):
    """Run a command and return result."""
    print(f"Running: {' '.join(cmd)}")
    try:
        result = subprocess.run(
            cmd, 
            cwd=cwd, 
            capture_output=capture_output, 
            text=True, 
            check=True
        )
        if result.stdout and capture_output:
            print(f"STDOUT: {result.stdout}")
        return True, result.stdout if capture_output else ""
    except subprocess.CalledProcessError as e:
        print(f"Error: {e}")
        if e.stdout:
            print(f"STDOUT: {e.stdout}")
        if e.stderr:
            print(f"STDERR: {e.stderr}")
        return False, ""


class InstallationTester:
    """Test package installation in various scenarios."""
    
    def __init__(self, package_path=None):
        self.package_path = package_path or "word-mcp-server"
        self.test_results = []
    
    def log_result(self, test_name, success, message=""):
        """Log test result."""
        status = "✅ PASS" if success else "❌ FAIL"
        print(f"{status}: {test_name}")
        if message:
            print(f"  {message}")
        
        self.test_results.append({
            "test": test_name,
            "success": success,
            "message": message
        })
    
    def test_pip_install(self, venv_path):
        """Test pip installation."""
        print("\n" + "="*50)
        print("Testing pip installation...")
        
        # Get pip path
        if os.name == 'nt':
            pip_path = venv_path / "Scripts" / "pip.exe"
            python_path = venv_path / "Scripts" / "python.exe"
        else:
            pip_path = venv_path / "bin" / "pip"
            python_path = venv_path / "bin" / "python"
        
        # Install package
        success, output = run_command([str(pip_path), "install", self.package_path])
        self.log_result("pip install", success, output if not success else "")
        
        if not success:
            return False
        
        # Verify installation
        success, output = run_command([str(python_path), "-c", "import word_mcp_server; print('Import successful')"])
        self.log_result("Package import", success, output if not success else "")
        
        return success
    
    def test_entry_points(self, venv_path):
        """Test entry point scripts."""
        print("\n" + "="*50)
        print("Testing entry points...")
        
        if os.name == 'nt':
            script_path = venv_path / "Scripts" / "word-mcp-server.exe"
        else:
            script_path = venv_path / "bin" / "word-mcp-server"
        
        # Test --version
        success, output = run_command([str(script_path), "--version"])
        self.log_result("Entry point --version", success, output if not success else "")
        
        # Test --help
        success, output = run_command([str(script_path), "--help"])
        self.log_result("Entry point --help", success, output if not success else "")
        
        # Test --check-requirements
        success, output = run_command([str(script_path), "--check-requirements"])
        self.log_result("Entry point --check-requirements", success, output if not success else "")
        
        return all([
            self.test_results[-3]["success"],
            self.test_results[-2]["success"],
            self.test_results[-1]["success"]
        ])
    
    def test_module_execution(self, venv_path):
        """Test module execution with python -m."""
        print("\n" + "="*50)
        print("Testing module execution...")
        
        if os.name == 'nt':
            python_path = venv_path / "Scripts" / "python.exe"
        else:
            python_path = venv_path / "bin" / "python"
        
        # Test python -m word_mcp_server --version
        success, output = run_command([str(python_path), "-m", "word_mcp_server", "--version"])
        self.log_result("Module execution --version", success, output if not success else "")
        
        # Test python -m word_mcp_server --help
        success, output = run_command([str(python_path), "-m", "word_mcp_server", "--help"])
        self.log_result("Module execution --help", success, output if not success else "")
        
        return all([
            self.test_results[-2]["success"],
            self.test_results[-1]["success"]
        ])
    
    def test_dependencies(self, venv_path):
        """Test that all dependencies are installed correctly."""
        print("\n" + "="*50)
        print("Testing dependencies...")
        
        if os.name == 'nt':
            python_path = venv_path / "Scripts" / "python.exe"
        else:
            python_path = venv_path / "bin" / "python"
        
        dependencies = [
            "mcp",
            "pydantic", 
            "yaml",
            "docx",
            "aiofiles"
        ]
        
        # Test Windows-specific dependencies only on Windows
        if os.name == 'nt':
            dependencies.extend(["win32com", "pythoncom"])
        
        all_success = True
        for dep in dependencies:
            if dep == "yaml":
                import_cmd = "import yaml"
            elif dep == "docx":
                import_cmd = "import docx"
            elif dep == "win32com":
                import_cmd = "import win32com.client"
            elif dep == "pythoncom":
                import_cmd = "import pythoncom"
            else:
                import_cmd = f"import {dep}"
            
            success, output = run_command([
                str(python_path), "-c", 
                f"{import_cmd}; print('{dep} imported successfully')"
            ])
            self.log_result(f"Dependency {dep}", success, output if not success else "")
            
            if not success:
                all_success = False
        
        return all_success
    
    def test_config_creation(self, venv_path):
        """Test configuration file creation."""
        print("\n" + "="*50)
        print("Testing configuration creation...")
        
        if os.name == 'nt':
            script_path = venv_path / "Scripts" / "word-mcp-server.exe"
        else:
            script_path = venv_path / "bin" / "word-mcp-server"
        
        # Create temporary directory for config
        with tempfile.TemporaryDirectory() as temp_dir:
            config_path = Path(temp_dir) / "test_config.yaml"
            
            # Test config creation
            success, output = run_command([
                str(script_path), "--create-config", "-o", str(config_path)
            ])
            self.log_result("Config creation", success, output if not success else "")
            
            if success:
                # Verify config file exists and is valid
                config_exists = config_path.exists()
                self.log_result("Config file exists", config_exists)
                
                if config_exists:
                    # Try to read the config
                    try:
                        import yaml
                        with open(config_path, 'r') as f:
                            config_data = yaml.safe_load(f)
                        
                        # Check for required sections
                        required_sections = ['server', 'word', 'logging', 'security']
                        sections_present = all(section in config_data for section in required_sections)
                        self.log_result("Config structure valid", sections_present)
                        
                        return success and config_exists and sections_present
                    except Exception as e:
                        self.log_result("Config parsing", False, str(e))
                        return False
            
            return success
    
    def run_all_tests(self):
        """Run all installation tests."""
        print("Word MCP Server Installation Test Suite")
        print("="*60)
        
        # Create temporary directory for testing
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            venv_path = temp_path / "test_venv"
            
            print(f"Creating test environment in: {venv_path}")
            
            # Create virtual environment
            try:
                venv.create(venv_path, with_pip=True)
                print("✅ Virtual environment created")
            except Exception as e:
                print(f"❌ Failed to create virtual environment: {e}")
                return False
            
            # Run tests
            tests = [
                ("Package Installation", lambda: self.test_pip_install(venv_path)),
                ("Entry Points", lambda: self.test_entry_points(venv_path)),
                ("Module Execution", lambda: self.test_module_execution(venv_path)),
                ("Dependencies", lambda: self.test_dependencies(venv_path)),
                ("Configuration", lambda: self.test_config_creation(venv_path)),
            ]
            
            overall_success = True
            for test_name, test_func in tests:
                try:
                    success = test_func()
                    if not success:
                        overall_success = False
                except Exception as e:
                    print(f"❌ {test_name} failed with exception: {e}")
                    overall_success = False
            
            # Print summary
            self.print_summary()
            
            return overall_success
    
    def print_summary(self):
        """Print test summary."""
        print("\n" + "="*60)
        print("TEST SUMMARY")
        print("="*60)
        
        passed = sum(1 for result in self.test_results if result["success"])
        total = len(self.test_results)
        
        print(f"Tests passed: {passed}/{total}")
        
        if passed == total:
            print("🎉 All tests passed! Package installation is working correctly.")
        else:
            print("❌ Some tests failed. Please review the issues above.")
            print("\nFailed tests:")
            for result in self.test_results:
                if not result["success"]:
                    print(f"  - {result['test']}: {result['message']}")


def main():
    """Main test runner."""
    import argparse
    
    parser = argparse.ArgumentParser(description="Test Word MCP Server installation")
    parser.add_argument(
        "--package", "-p",
        help="Package to install (default: word-mcp-server from PyPI)",
        default="word-mcp-server"
    )
    parser.add_argument(
        "--local", "-l",
        action="store_true",
        help="Test local package (use current directory)"
    )
    
    args = parser.parse_args()
    
    if args.local:
        # Test local package
        package_path = "."
    else:
        package_path = args.package
    
    tester = InstallationTester(package_path)
    success = tester.run_all_tests()
    
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()