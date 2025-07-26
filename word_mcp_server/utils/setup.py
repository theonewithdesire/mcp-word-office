"""
Setup and installation utilities for Word MCP Server.
"""

import os
import sys
import json
import shutil
import platform
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from ..config.config_manager import ConfigManager


class SetupManager:
    """Manages installation and setup of Word MCP Server."""
    
    def __init__(self):
        self.platform = platform.system().lower()
        self.is_windows = self.platform == "windows"
        self.home_dir = Path.home()
        self.config_dir = self.home_dir / ".word_mcp_server"
        
    def check_system_requirements(self) -> Tuple[bool, List[str]]:
        """Check if system meets requirements for Word MCP Server.
        
        Returns:
            Tuple of (requirements_met, issues_list)
        """
        issues = []
        
        # Check Python version
        if sys.version_info < (3, 8):
            issues.append(f"Python 3.8+ required, found {sys.version}")
        
        # Check if Windows (required for Word COM)
        if not self.is_windows:
            issues.append("Microsoft Word COM automation requires Windows")
        
        # Check for Word installation (Windows only)
        if self.is_windows:
            word_installed = self._check_word_installation()
            if not word_installed:
                issues.append("Microsoft Word not found - please install Microsoft Office")
        
        # Check required Python packages
        missing_packages = self._check_python_packages()
        if missing_packages:
            issues.append(f"Missing Python packages: {', '.join(missing_packages)}")
        
        return len(issues) == 0, issues
    
    def _check_word_installation(self) -> bool:
        """Check if Microsoft Word is installed (Windows only)."""
        if not self.is_windows:
            return False
        
        try:
            import winreg
            
            # Check common registry locations for Word
            registry_paths = [
                r"SOFTWARE\Microsoft\Office\16.0\Word\InstallRoot",  # Office 2016/2019/365
                r"SOFTWARE\Microsoft\Office\15.0\Word\InstallRoot",  # Office 2013
                r"SOFTWARE\Microsoft\Office\14.0\Word\InstallRoot",  # Office 2010
            ]
            
            for path in registry_paths:
                try:
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path):
                        return True
                except FileNotFoundError:
                    continue
            
            # Also check if winword.exe is in PATH
            word_exe = shutil.which("winword.exe")
            return word_exe is not None
            
        except ImportError:
            # winreg not available (not Windows)
            return False
        except Exception:
            return False
    
    def _check_python_packages(self) -> List[str]:
        """Check for required Python packages."""
        required_packages = [
            "mcp",
            "pydantic", 
            "pyyaml",
            "python-docx"
        ]
        
        if self.is_windows:
            required_packages.append("pywin32")
        
        missing = []
        for package in required_packages:
            try:
                __import__(package.replace("-", "_"))
            except ImportError:
                missing.append(package)
        
        return missing
    
    def create_config_directory(self) -> None:
        """Create configuration directory structure."""
        directories = [
            self.config_dir,
            self.config_dir / "logs",
            self.config_dir / "backups",
            self.config_dir / "temp"
        ]
        
        for directory in directories:
            directory.mkdir(parents=True, exist_ok=True)
            print(f"✓ Created directory: {directory}")
    
    def install_default_config(self) -> str:
        """Install default configuration file.
        
        Returns:
            Path to created configuration file
        """
        config_path = self.config_dir / "config.yaml"
        
        if config_path.exists():
            backup_path = config_path.with_suffix(".yaml.backup")
            shutil.copy2(config_path, backup_path)
            print(f"✓ Backed up existing config to: {backup_path}")
        
        config_manager = ConfigManager()
        config_manager.create_default_config(str(config_path))
        
        print(f"✓ Created default configuration: {config_path}")
        return str(config_path)
    
    def setup_claude_integration(self, config_path: Optional[str] = None) -> Dict[str, str]:
        """Generate Claude MCP configuration.
        
        Args:
            config_path: Optional path to Word MCP Server config
            
        Returns:
            Dictionary with Claude MCP configuration
        """
        if config_path is None:
            config_path = str(self.config_dir / "config.yaml")
        
        # Generate Claude MCP server configuration
        claude_config = {
            "mcpServers": {
                "word-office": {
                    "command": "python",
                    "args": ["-m", "word_mcp_server"],
                    "env": {
                        "WORD_MCP_CONFIG": config_path
                    }
                }
            }
        }
        
        # Save Claude configuration
        claude_config_path = self.config_dir / "claude_mcp_config.json"
        with open(claude_config_path, 'w') as f:
            json.dump(claude_config, f, indent=2)
        
        print(f"✓ Generated Claude MCP configuration: {claude_config_path}")
        
        return claude_config
    
    def create_startup_scripts(self) -> List[str]:
        """Create startup scripts for different platforms.
        
        Returns:
            List of created script paths
        """
        scripts = []
        
        if self.is_windows:
            # Create Windows batch script
            batch_script = self.config_dir / "start_word_mcp_server.bat"
            batch_content = f"""@echo off
echo Starting Word MCP Server...
python -m word_mcp_server --config "{self.config_dir / 'config.yaml'}"
pause
"""
            with open(batch_script, 'w') as f:
                f.write(batch_content)
            scripts.append(str(batch_script))
            print(f"✓ Created Windows startup script: {batch_script}")
        
        # Create cross-platform shell script
        shell_script = self.config_dir / "start_word_mcp_server.sh"
        shell_content = f"""#!/bin/bash
echo "Starting Word MCP Server..."
python -m word_mcp_server --config "{self.config_dir / 'config.yaml'}"
"""
        with open(shell_script, 'w') as f:
            f.write(shell_content)
        
        # Make shell script executable
        os.chmod(shell_script, 0o755)
        scripts.append(str(shell_script))
        print(f"✓ Created shell startup script: {shell_script}")
        
        return scripts
    
    def run_full_setup(self) -> Dict[str, any]:
        """Run complete setup process with enhanced error handling and recovery.
        
        Returns:
            Dictionary with setup results
        """
        print("Word MCP Server Setup")
        print("=" * 50)
        
        results = {
            "success": True,
            "requirements_met": False,
            "issues": [],
            "config_path": None,
            "claude_config": None,
            "startup_scripts": [],
            "setup_phase": "initialization"
        }
        
        try:
            # Phase 1: Check system requirements
            results["setup_phase"] = "requirements_check"
            print("\n1. Checking system requirements...")
            requirements_met, issues = self.check_system_requirements()
            results["requirements_met"] = requirements_met
            results["issues"] = issues
            
            if not requirements_met:
                print("✗ System requirements not met:")
                for issue in issues:
                    print(f"  - {issue}")
                
                # Provide helpful suggestions for common issues
                self._provide_requirement_suggestions(issues)
                results["success"] = False
                return results
            
            print("✓ System requirements met")
            
            # Phase 2: Create configuration directory
            results["setup_phase"] = "directory_creation"
            print("\n2. Creating configuration directory...")
            try:
                self.create_config_directory()
            except PermissionError as e:
                print(f"✗ Permission denied creating configuration directory: {e}")
                print("  Try running as administrator or check directory permissions")
                results["success"] = False
                return results
            except Exception as e:
                print(f"✗ Failed to create configuration directory: {e}")
                results["success"] = False
                return results
            
            # Phase 3: Install default configuration
            results["setup_phase"] = "configuration_install"
            print("\n3. Installing default configuration...")
            try:
                config_path = self.install_default_config()
                results["config_path"] = config_path
                
                # Validate the created configuration
                self._validate_config_file(config_path)
                print("✓ Configuration file validated")
                
            except Exception as e:
                print(f"✗ Failed to install configuration: {e}")
                results["success"] = False
                return results
            
            # Phase 4: Setup Claude integration
            results["setup_phase"] = "claude_integration"
            print("\n4. Setting up Claude integration...")
            try:
                claude_config = self.setup_claude_integration(config_path)
                results["claude_config"] = claude_config
                
                # Provide Claude setup instructions
                self._provide_claude_setup_instructions()
                
            except Exception as e:
                print(f"✗ Failed to setup Claude integration: {e}")
                results["success"] = False
                return results
            
            # Phase 5: Create startup scripts
            results["setup_phase"] = "startup_scripts"
            print("\n5. Creating startup scripts...")
            try:
                startup_scripts = self.create_startup_scripts()
                results["startup_scripts"] = startup_scripts
                
                # Make scripts executable and test them
                self._validate_startup_scripts(startup_scripts)
                
            except Exception as e:
                print(f"✗ Failed to create startup scripts: {e}")
                results["success"] = False
                return results
            
            # Phase 6: Final validation and testing
            results["setup_phase"] = "final_validation"
            print("\n6. Running final validation...")
            try:
                validation_results = self._run_final_validation(config_path)
                if not validation_results["success"]:
                    print("⚠ Setup completed with warnings:")
                    for warning in validation_results["warnings"]:
                        print(f"  - {warning}")
                else:
                    print("✓ All validation checks passed")
                    
            except Exception as e:
                print(f"⚠ Final validation failed: {e}")
                # Don't fail setup for validation issues
            
            # Success summary
            results["setup_phase"] = "completed"
            print("\n" + "=" * 50)
            print("✓ Setup completed successfully!")
            print(f"\nConfiguration directory: {self.config_dir}")
            print(f"Configuration file: {config_path}")
            print(f"Claude MCP config: {self.config_dir / 'claude_mcp_config.json'}")
            
            print("\nNext steps:")
            print("1. Review and customize the configuration file if needed")
            print("2. Add the Claude MCP configuration to your Claude settings")
            print("3. Start the server using one of the startup scripts")
            print("4. Test the connection with Claude")
            
            # Provide platform-specific instructions
            self._provide_platform_specific_instructions()
            
            return results
            
        except KeyboardInterrupt:
            print("\n✗ Setup interrupted by user")
            results["success"] = False
            results["setup_phase"] = "interrupted"
            return results
        except Exception as e:
            print(f"\n✗ Unexpected error during {results['setup_phase']}: {e}")
            results["success"] = False
            return results
    
    def _provide_requirement_suggestions(self, issues: List[str]) -> None:
        """Provide helpful suggestions for requirement issues."""
        print("\nSuggestions to resolve issues:")
        
        for issue in issues:
            if "Python" in issue and "required" in issue:
                print("  • Update Python: Visit https://python.org/downloads/")
            elif "Windows" in issue:
                print("  • Word MCP Server requires Windows for COM automation")
                print("  • Consider using Windows Subsystem for Linux (WSL) with Windows")
            elif "Word not found" in issue:
                print("  • Install Microsoft Office from https://office.microsoft.com/")
                print("  • Ensure Word is properly registered in Windows")
            elif "Missing Python packages" in issue:
                print("  • Install missing packages: pip install word-mcp-server[all]")
                print("  • Or install individually: pip install mcp pydantic pyyaml python-docx pywin32")
    
    def _validate_config_file(self, config_path: str) -> None:
        """Validate the created configuration file."""
        from ..config.config_manager import ConfigManager
        
        try:
            # Try to load the configuration to ensure it's valid
            config_manager = ConfigManager(config_path)
            config = config_manager.load_config()
            
            # Basic validation checks
            assert config.server.port > 0, "Invalid server port"
            assert config.server.max_concurrent_docs > 0, "Invalid max concurrent docs"
            assert config.logging.level in ["DEBUG", "INFO", "WARNING", "ERROR"], "Invalid log level"
            
        except Exception as e:
            raise Exception(f"Configuration validation failed: {e}")
    
    def _provide_claude_setup_instructions(self) -> None:
        """Provide detailed Claude setup instructions."""
        claude_config_path = self.config_dir / "claude_mcp_config.json"
        
        print(f"✓ Claude configuration saved to: {claude_config_path}")
        print("\nTo add this to Claude:")
        print("1. Open Claude Desktop application")
        print("2. Go to Settings > Developer")
        print("3. Edit the MCP servers configuration")
        print("4. Add the contents of claude_mcp_config.json to your configuration")
        print("5. Restart Claude Desktop")
    
    def _validate_startup_scripts(self, startup_scripts: List[str]) -> None:
        """Validate and test startup scripts."""
        for script_path in startup_scripts:
            script_file = Path(script_path)
            
            if not script_file.exists():
                raise Exception(f"Startup script not created: {script_path}")
            
            # Check if script is executable (Unix-like systems)
            if script_file.suffix == '.sh':
                stat = script_file.stat()
                if not (stat.st_mode & 0o111):
                    raise Exception(f"Startup script not executable: {script_path}")
            
            # Basic content validation
            with open(script_path, 'r') as f:
                content = f.read()
                if 'word_mcp_server' not in content:
                    raise Exception(f"Invalid startup script content: {script_path}")
    
    def _run_final_validation(self, config_path: str) -> Dict[str, any]:
        """Run final validation checks."""
        validation_results = {
            "success": True,
            "warnings": []
        }
        
        try:
            # Test configuration loading
            from ..config.config_manager import ConfigManager
            config_manager = ConfigManager(config_path)
            config = config_manager.load_config()
            
            # Check if Word is accessible
            if self.is_windows:
                try:
                    import win32com.client
                    word_app = win32com.client.Dispatch("Word.Application")
                    word_app.Quit()
                    print("✓ Word COM interface accessible")
                except Exception as e:
                    validation_results["warnings"].append(f"Word COM interface test failed: {e}")
            
            # Check directory permissions
            test_file = self.config_dir / "test_write.tmp"
            try:
                test_file.write_text("test")
                test_file.unlink()
                print("✓ Configuration directory writable")
            except Exception as e:
                validation_results["warnings"].append(f"Configuration directory not writable: {e}")
            
            # Check if all required files exist
            required_files = [
                self.config_dir / "config.yaml",
                self.config_dir / "claude_mcp_config.json"
            ]
            
            for required_file in required_files:
                if not required_file.exists():
                    validation_results["warnings"].append(f"Required file missing: {required_file}")
            
            if validation_results["warnings"]:
                validation_results["success"] = False
            
        except Exception as e:
            validation_results["success"] = False
            validation_results["warnings"].append(f"Validation error: {e}")
        
        return validation_results
    
    def _provide_platform_specific_instructions(self) -> None:
        """Provide platform-specific setup instructions."""
        print(f"\nPlatform-specific notes ({self.platform}):")
        
        if self.is_windows:
            print("• Windows detected - full Word COM automation available")
            print("• You can run the .bat file to start the server")
            print("• Consider adding the server to Windows startup if needed")
        else:
            print("• Non-Windows platform detected")
            print("• Word COM automation requires Windows")
            print("• Consider using Windows VM or WSL for full functionality")
    
    def uninstall(self) -> bool:
        """Uninstall Word MCP Server configuration and files.
        
        Returns:
            True if successful, False otherwise
        """
        try:
            if self.config_dir.exists():
                # Create backup before removal
                backup_dir = self.config_dir.with_suffix(".backup")
                if backup_dir.exists():
                    shutil.rmtree(backup_dir)
                
                shutil.move(str(self.config_dir), str(backup_dir))
                print(f"✓ Configuration moved to backup: {backup_dir}")
                print("✓ Word MCP Server uninstalled successfully")
                return True
            else:
                print("✓ No configuration found, nothing to uninstall")
                return True
                
        except Exception as e:
            print(f"✗ Failed to uninstall: {e}")
            return False


def main():
    """Main setup entry point."""
    import argparse
    
    parser = argparse.ArgumentParser(description="Word MCP Server Setup Utility")
    parser.add_argument("--check", action="store_true", help="Check system requirements only")
    parser.add_argument("--uninstall", action="store_true", help="Uninstall Word MCP Server")
    parser.add_argument("--config-only", action="store_true", help="Create configuration only")
    
    args = parser.parse_args()
    
    setup_manager = SetupManager()
    
    if args.uninstall:
        setup_manager.uninstall()
        return
    
    if args.check:
        requirements_met, issues = setup_manager.check_system_requirements()
        if requirements_met:
            print("✓ All system requirements met")
        else:
            print("✗ System requirements not met:")
            for issue in issues:
                print(f"  - {issue}")
        return
    
    if args.config_only:
        setup_manager.create_config_directory()
        setup_manager.install_default_config()
        return
    
    # Run full setup
    setup_manager.run_full_setup()


if __name__ == "__main__":
    main()