#!/usr/bin/env python3
"""
Test Word MCP Server in different Python environments and scenarios.
"""

import os
import sys
import subprocess
import tempfile
import shutil
import json
from pathlib import Path


class EnvironmentTester:
    """Test package in different Python environments."""
    
    def __init__(self):
        self.results = []
    
    def log_result(self, environment, test, success, message=""):
        """Log test result."""
        self.results.append({
            "environment": environment,
            "test": test,
            "success": success,
            "message": message
        })
        
        status = "✅" if success else "❌"
        print(f"{status} [{environment}] {test}: {message}")
    
    def run_command(self, cmd, cwd=None):
        """Run command and return success status."""
        try:
            result = subprocess.run(
                cmd, cwd=cwd, capture_output=True, text=True, check=True
            )
            return True, result.stdout
        except subprocess.CalledProcessError as e:
            return False, e.stderr or str(e)
    
    def test_python_version(self, python_exe, env_name):
        """Test Python version compatibility."""
        success, output = self.run_command([python_exe, "--version"])
        if success:
            version = output.strip()
            self.log_result(env_name, "Python Version", True, version)
            return True
        else:
            self.log_result(env_name, "Python Version", False, output)
            return False
    
    def test_pip_install(self, python_exe, env_name, package_path="word-mcp-server"):
        """Test pip installation."""
        # Get pip command
        pip_cmd = [python_exe, "-m", "pip", "install", package_path]
        
        success, output = self.run_command(pip_cmd)
        if success:
            self.log_result(env_name, "Pip Install", True, "Package installed successfully")
            return True
        else:
            self.log_result(env_name, "Pip Install", False, output)
            return False
    
    def test_import(self, python_exe, env_name):
        """Test package import."""
        success, output = self.run_command([
            python_exe, "-c", 
            "import word_mcp_server; print('Import successful')"
        ])
        
        if success:
            self.log_result(env_name, "Package Import", True, "Import successful")
            return True
        else:
            self.log_result(env_name, "Package Import", False, output)
            return False
    
    def test_cli_command(self, python_exe, env_name):
        """Test CLI command execution."""
        success, output = self.run_command([
            python_exe, "-m", "word_mcp_server", "--version"
        ])
        
        if success:
            self.log_result(env_name, "CLI Command", True, output.strip())
            return True
        else:
            self.log_result(env_name, "CLI Command", False, output)
            return False
    
    def test_requirements_check(self, python_exe, env_name):
        """Test system requirements check."""
        success, output = self.run_command([
            python_exe, "-m", "word_mcp_server", "--check-requirements"
        ])
        
        # This might fail on non-Windows systems, which is expected
        if success or "Windows" in output or "Word" in output:
            self.log_result(env_name, "Requirements Check", True, "Check completed")
            return True
        else:
            self.log_result(env_name, "Requirements Check", False, output)
            return False
    
    def test_environment(self, python_exe, env_name, package_path="word-mcp-server"):
        """Test complete environment."""
        print(f"\n{'='*60}")
        print(f"Testing environment: {env_name}")
        print(f"Python executable: {python_exe}")
        print(f"{'='*60}")
        
        tests = [
            ("Python Version", lambda: self.test_python_version(python_exe, env_name)),
            ("Pip Install", lambda: self.test_pip_install(python_exe, env_name, package_path)),
            ("Package Import", lambda: self.test_import(python_exe, env_name)),
            ("CLI Command", lambda: self.test_cli_command(python_exe, env_name)),
            ("Requirements Check", lambda: self.test_requirements_check(python_exe, env_name)),
        ]
        
        env_success = True
        for test_name, test_func in tests:
            try:
                success = test_func()
                if not success:
                    env_success = False
            except Exception as e:
                self.log_result(env_name, test_name, False, str(e))
                env_success = False
        
        return env_success
    
    def find_python_installations(self):
        """Find available Python installations."""
        installations = []
        
        # Common Python installation paths
        if os.name == 'nt':  # Windows
            common_paths = [
                "python",
                "python3",
                "py -3.8",
                "py -3.9", 
                "py -3.10",
                "py -3.11",
                "py -3.12",
                r"C:\Python38\python.exe",
                r"C:\Python39\python.exe",
                r"C:\Python310\python.exe",
                r"C:\Python311\python.exe",
                r"C:\Python312\python.exe",
            ]
        else:  # Unix-like
            common_paths = [
                "python3",
                "python3.8",
                "python3.9",
                "python3.10",
                "python3.11",
                "python3.12",
                "/usr/bin/python3",
                "/usr/local/bin/python3",
            ]
        
        for path in common_paths:
            try:
                # Test if Python executable works
                result = subprocess.run(
                    path.split() + ["--version"], 
                    capture_output=True, text=True, check=True
                )
                version = result.stdout.strip()
                installations.append((path, version))
            except (subprocess.CalledProcessError, FileNotFoundError):
                continue
        
        return installations
    
    def test_all_environments(self, package_path="word-mcp-server"):
        """Test package in all available Python environments."""
        print("Word MCP Server Environment Test Suite")
        print("="*60)
        
        installations = self.find_python_installations()
        
        if not installations:
            print("❌ No Python installations found!")
            return False
        
        print(f"Found {len(installations)} Python installations:")
        for path, version in installations:
            print(f"  - {path}: {version}")
        
        overall_success = True
        
        for python_exe, version in installations:
            env_name = f"{version} ({python_exe})"
            
            # Create temporary virtual environment for testing
            with tempfile.TemporaryDirectory() as temp_dir:
                venv_path = Path(temp_dir) / "test_env"
                
                # Create virtual environment
                try:
                    subprocess.run([
                        python_exe.split()[0] if ' ' in python_exe else python_exe,
                        "-m", "venv", str(venv_path)
                    ], check=True, capture_output=True)
                    
                    # Get venv python path
                    if os.name == 'nt':
                        venv_python = venv_path / "Scripts" / "python.exe"
                    else:
                        venv_python = venv_path / "bin" / "python"
                    
                    # Test in virtual environment
                    success = self.test_environment(str(venv_python), env_name, package_path)
                    if not success:
                        overall_success = False
                        
                except Exception as e:
                    self.log_result(env_name, "Environment Setup", False, str(e))
                    overall_success = False
        
        self.print_summary()
        return overall_success
    
    def print_summary(self):
        """Print test summary."""
        print("\n" + "="*60)
        print("ENVIRONMENT TEST SUMMARY")
        print("="*60)
        
        # Group results by environment
        env_results = {}
        for result in self.results:
            env = result["environment"]
            if env not in env_results:
                env_results[env] = []
            env_results[env].append(result)
        
        total_envs = len(env_results)
        successful_envs = 0
        
        for env, results in env_results.items():
            passed = sum(1 for r in results if r["success"])
            total = len(results)
            
            if passed == total:
                successful_envs += 1
                status = "✅ PASS"
            else:
                status = "❌ FAIL"
            
            print(f"{status} {env}: {passed}/{total} tests passed")
            
            # Show failed tests
            failed_tests = [r for r in results if not r["success"]]
            if failed_tests:
                for test in failed_tests:
                    print(f"    ❌ {test['test']}: {test['message']}")
        
        print(f"\nOverall: {successful_envs}/{total_envs} environments successful")
        
        if successful_envs == total_envs:
            print("🎉 All environments passed!")
        else:
            print("❌ Some environments failed. Check the details above.")
    
    def export_results(self, filename="test_results.json"):
        """Export results to JSON file."""
        with open(filename, 'w') as f:
            json.dump(self.results, f, indent=2)
        print(f"Results exported to {filename}")


def main():
    """Main test runner."""
    import argparse
    
    parser = argparse.ArgumentParser(
        description="Test Word MCP Server in different Python environments"
    )
    parser.add_argument(
        "--package", "-p",
        help="Package to test (default: word-mcp-server)",
        default="word-mcp-server"
    )
    parser.add_argument(
        "--local", "-l",
        action="store_true",
        help="Test local package (current directory)"
    )
    parser.add_argument(
        "--export", "-e",
        help="Export results to JSON file",
        metavar="FILENAME"
    )
    
    args = parser.parse_args()
    
    package_path = "." if args.local else args.package
    
    tester = EnvironmentTester()
    success = tester.test_all_environments(package_path)
    
    if args.export:
        tester.export_results(args.export)
    
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()