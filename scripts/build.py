#!/usr/bin/env python3
"""
Build script for Word MCP Server distribution.
"""

import os
import sys
import shutil
import subprocess
from pathlib import Path


def run_command(cmd, cwd=None):
    """Run a command and return success status."""
    print(f"Running: {' '.join(cmd)}")
    try:
        result = subprocess.run(cmd, cwd=cwd, check=True, capture_output=True, text=True)
        if result.stdout:
            print(result.stdout)
        return True
    except subprocess.CalledProcessError as e:
        print(f"Error: {e}")
        if e.stdout:
            print(f"STDOUT: {e.stdout}")
        if e.stderr:
            print(f"STDERR: {e.stderr}")
        return False


def clean_build():
    """Clean previous build artifacts."""
    print("Cleaning build artifacts...")
    
    dirs_to_clean = ["build", "dist", "*.egg-info"]
    for pattern in dirs_to_clean:
        for path in Path(".").glob(pattern):
            if path.is_dir():
                print(f"Removing directory: {path}")
                shutil.rmtree(path)
            elif path.is_file():
                print(f"Removing file: {path}")
                path.unlink()
    
    # Clean __pycache__ directories
    for pycache in Path(".").rglob("__pycache__"):
        print(f"Removing __pycache__: {pycache}")
        shutil.rmtree(pycache)
    
    # Clean .pyc files
    for pyc in Path(".").rglob("*.pyc"):
        print(f"Removing .pyc file: {pyc}")
        pyc.unlink()


def run_tests():
    """Run the test suite."""
    print("Running tests...")
    return run_command([sys.executable, "-m", "pytest", "tests/", "-v"])


def run_linting():
    """Run code quality checks."""
    print("Running code quality checks...")
    
    # Check if tools are available
    tools = ["black", "isort", "flake8", "mypy"]
    available_tools = []
    
    for tool in tools:
        try:
            subprocess.run([tool, "--version"], capture_output=True, check=True)
            available_tools.append(tool)
        except (subprocess.CalledProcessError, FileNotFoundError):
            print(f"Warning: {tool} not available, skipping...")
    
    success = True
    
    if "black" in available_tools:
        print("Checking code formatting with black...")
        if not run_command([sys.executable, "-m", "black", "--check", "word_mcp_server/"]):
            success = False
    
    if "isort" in available_tools:
        print("Checking import sorting with isort...")
        if not run_command([sys.executable, "-m", "isort", "--check-only", "word_mcp_server/"]):
            success = False
    
    if "flake8" in available_tools:
        print("Running flake8 linting...")
        if not run_command([sys.executable, "-m", "flake8", "word_mcp_server/"]):
            success = False
    
    if "mypy" in available_tools:
        print("Running mypy type checking...")
        if not run_command([sys.executable, "-m", "mypy", "word_mcp_server/"]):
            success = False
    
    return success


def build_package():
    """Build the distribution packages."""
    print("Building distribution packages...")
    
    # Build source distribution and wheel
    if not run_command([sys.executable, "-m", "build"]):
        return False
    
    print("Build completed successfully!")
    print("Distribution files created in dist/:")
    
    dist_dir = Path("dist")
    if dist_dir.exists():
        for file in dist_dir.iterdir():
            print(f"  - {file.name}")
    
    return True


def verify_package():
    """Verify the built package."""
    print("Verifying package...")
    
    # Check if twine is available for verification
    try:
        subprocess.run(["twine", "--version"], capture_output=True, check=True)
        return run_command(["twine", "check", "dist/*"])
    except (subprocess.CalledProcessError, FileNotFoundError):
        print("Warning: twine not available, skipping package verification")
        return True


def main():
    """Main build process."""
    print("Word MCP Server Build Script")
    print("=" * 40)
    
    # Change to project root
    script_dir = Path(__file__).parent
    project_root = script_dir.parent
    os.chdir(project_root)
    
    steps = [
        ("Clean build artifacts", clean_build),
        ("Run tests", run_tests),
        ("Run linting", run_linting),
        ("Build package", build_package),
        ("Verify package", verify_package),
    ]
    
    for step_name, step_func in steps:
        print(f"\n{step_name}...")
        print("-" * 30)
        
        if not step_func():
            print(f"❌ {step_name} failed!")
            sys.exit(1)
        
        print(f"✅ {step_name} completed successfully!")
    
    print("\n" + "=" * 40)
    print("🎉 Build process completed successfully!")
    print("\nNext steps:")
    print("1. Test the package: pip install dist/word_mcp_server-*.whl")
    print("2. Upload to PyPI: twine upload dist/*")


if __name__ == "__main__":
    main()