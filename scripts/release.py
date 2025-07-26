#!/usr/bin/env python3
"""
Release script for Word MCP Server.
"""

import os
import sys
import re
import subprocess
from pathlib import Path
from datetime import datetime


def run_command(cmd, cwd=None):
    """Run a command and return success status."""
    print(f"Running: {' '.join(cmd)}")
    try:
        result = subprocess.run(cmd, cwd=cwd, check=True, capture_output=True, text=True)
        if result.stdout:
            print(result.stdout)
        return True, result.stdout
    except subprocess.CalledProcessError as e:
        print(f"Error: {e}")
        if e.stdout:
            print(f"STDOUT: {e.stdout}")
        if e.stderr:
            print(f"STDERR: {e.stderr}")
        return False, ""


def get_current_version():
    """Get current version from pyproject.toml."""
    pyproject_path = Path("pyproject.toml")
    if not pyproject_path.exists():
        print("Error: pyproject.toml not found")
        return None
    
    content = pyproject_path.read_text()
    match = re.search(r'version = "([^"]+)"', content)
    if match:
        return match.group(1)
    
    print("Error: Could not find version in pyproject.toml")
    return None


def update_version(new_version):
    """Update version in pyproject.toml."""
    pyproject_path = Path("pyproject.toml")
    content = pyproject_path.read_text()
    
    # Update version
    content = re.sub(r'version = "[^"]+"', f'version = "{new_version}"', content)
    pyproject_path.write_text(content)
    
    # Update version in main.py
    main_path = Path("word_mcp_server/main.py")
    if main_path.exists():
        main_content = main_path.read_text()
        main_content = re.sub(
            r'version="[^"]+"',
            f'version="%(prog)s {new_version}"',
            main_content
        )
        main_path.write_text(main_content)
    
    print(f"Updated version to {new_version}")


def update_changelog(version):
    """Update CHANGELOG.md with release information."""
    changelog_path = Path("CHANGELOG.md")
    if not changelog_path.exists():
        print("Warning: CHANGELOG.md not found")
        return
    
    content = changelog_path.read_text()
    today = datetime.now().strftime("%Y-%m-%d")
    
    # Replace [Unreleased] with version and date
    content = content.replace(
        "## [Unreleased]",
        f"## [Unreleased]\n\n### Added\n- TBD\n\n### Changed\n- TBD\n\n### Fixed\n- TBD\n\n## [{version}] - {today}"
    )
    
    changelog_path.write_text(content)
    print(f"Updated CHANGELOG.md for version {version}")


def create_git_tag(version):
    """Create and push git tag."""
    tag_name = f"v{version}"
    
    # Create tag
    success, _ = run_command(["git", "tag", "-a", tag_name, "-m", f"Release {version}"])
    if not success:
        return False
    
    # Push tag
    success, _ = run_command(["git", "push", "origin", tag_name])
    return success


def upload_to_pypi(test=False):
    """Upload package to PyPI."""
    if test:
        print("Uploading to Test PyPI...")
        return run_command(["twine", "upload", "--repository", "testpypi", "dist/*"])[0]
    else:
        print("Uploading to PyPI...")
        return run_command(["twine", "upload", "dist/*"])[0]


def main():
    """Main release process."""
    print("Word MCP Server Release Script")
    print("=" * 40)
    
    # Change to project root
    script_dir = Path(__file__).parent
    project_root = script_dir.parent
    os.chdir(project_root)
    
    # Get current version
    current_version = get_current_version()
    if not current_version:
        sys.exit(1)
    
    print(f"Current version: {current_version}")
    
    # Get new version from user
    new_version = input("Enter new version (or press Enter to keep current): ").strip()
    if not new_version:
        new_version = current_version
    
    # Validate version format
    if not re.match(r'^\d+\.\d+\.\d+', new_version):
        print("Error: Version must be in format X.Y.Z")
        sys.exit(1)
    
    # Ask for confirmation
    print(f"\nRelease summary:")
    print(f"  Version: {new_version}")
    print(f"  Current directory: {os.getcwd()}")
    
    confirm = input("\nProceed with release? (y/N): ").strip().lower()
    if confirm != 'y':
        print("Release cancelled")
        sys.exit(0)
    
    # Update version if changed
    if new_version != current_version:
        update_version(new_version)
        update_changelog(new_version)
    
    # Run build script
    print("\nRunning build process...")
    success, _ = run_command([sys.executable, "scripts/build.py"])
    if not success:
        print("Build failed, aborting release")
        sys.exit(1)
    
    # Ask about PyPI upload
    upload_choice = input("\nUpload to PyPI? (y/N/t for test): ").strip().lower()
    
    if upload_choice == 't':
        if not upload_to_pypi(test=True):
            print("Test PyPI upload failed")
            sys.exit(1)
        print("✅ Uploaded to Test PyPI successfully!")
    elif upload_choice == 'y':
        if not upload_to_pypi(test=False):
            print("PyPI upload failed")
            sys.exit(1)
        print("✅ Uploaded to PyPI successfully!")
    
    # Create git tag if version changed
    if new_version != current_version:
        tag_choice = input("\nCreate git tag? (y/N): ").strip().lower()
        if tag_choice == 'y':
            if create_git_tag(new_version):
                print("✅ Git tag created and pushed!")
            else:
                print("❌ Git tag creation failed")
    
    print("\n" + "=" * 40)
    print("🎉 Release process completed!")
    print(f"\nVersion {new_version} has been released.")
    
    if upload_choice in ['y', 't']:
        print("\nInstallation command:")
        if upload_choice == 't':
            print("pip install --index-url https://test.pypi.org/simple/ word-mcp-server")
        else:
            print("pip install word-mcp-server")


if __name__ == "__main__":
    main()