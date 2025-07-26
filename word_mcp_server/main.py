"""
Main entry point for Word MCP Server.
"""

import sys
import signal
import argparse
import asyncio
import logging
from pathlib import Path
from typing import Optional

from .config import ConfigManager
from .utils import setup_logging
from .utils.errors import WordMCPError


def parse_arguments() -> argparse.Namespace:
    """Parse command line arguments.
    
    Returns:
        Parsed arguments
    """
    parser = argparse.ArgumentParser(
        description="Word Office MCP Server",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  word-mcp-server                          # Start with default config
  word-mcp-server --config config.yaml    # Start with specific config
  word-mcp-server --create-config         # Create default config file
        """
    )
    
    parser.add_argument(
        "--config", "-c",
        type=str,
        help="Path to configuration file"
    )
    
    parser.add_argument(
        "--create-config",
        action="store_true",
        help="Create a default configuration file and exit"
    )
    
    parser.add_argument(
        "--config-output", "-o",
        type=str,
        default="config.yaml",
        help="Output path for created configuration file (default: config.yaml)"
    )
    
    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="Enable verbose logging"
    )
    
    parser.add_argument(
        "--setup",
        action="store_true",
        help="Run setup wizard to configure Word MCP Server"
    )
    
    parser.add_argument(
        "--check-requirements",
        action="store_true",
        help="Check system requirements"
    )
    
    parser.add_argument(
        "--uninstall",
        action="store_true",
        help="Uninstall Word MCP Server configuration"
    )
    
    parser.add_argument(
        "--version",
        action="version",
        version="%(prog)s 0.1.0"
    )
    
    return parser.parse_args()


def create_default_config(output_path: str) -> None:
    """Create a default configuration file.
    
    Args:
        output_path: Path where to create the config file
    """
    try:
        config_manager = ConfigManager()
        config = config_manager.create_default_config(output_path)
        print(f"✓ Default configuration created at: {output_path}")
        print("\nConfiguration summary:")
        print(f"  Server: {config.server.host}:{config.server.port}")
        print(f"  Max concurrent docs: {config.server.max_concurrent_docs}")
        print(f"  Logging level: {config.logging.level}")
        print(f"  Word auto-launch: {config.word.auto_launch}")
        print(f"  Backup enabled: {config.word.backup_enabled}")
        print(f"  Allowed paths: {len(config.security.allowed_paths)} configured")
        
    except Exception as e:
        print(f"✗ Failed to create configuration file: {e}", file=sys.stderr)
        sys.exit(1)


class ServerManager:
    """Manages server lifecycle and graceful shutdown."""
    
    def __init__(self):
        self.mcp_server: Optional['WordMCPServer'] = None
        self.logger: Optional[logging.Logger] = None
        self.shutdown_event = asyncio.Event()
        
    async def setup_signal_handlers(self):
        """Setup signal handlers for graceful shutdown."""
        def signal_handler(signum, frame):
            self.logger.info(f"Received signal {signum}, initiating graceful shutdown...")
            asyncio.create_task(self.shutdown())
        
        # Setup signal handlers for Unix systems
        if hasattr(signal, 'SIGTERM'):
            signal.signal(signal.SIGTERM, signal_handler)
        if hasattr(signal, 'SIGINT'):
            signal.signal(signal.SIGINT, signal_handler)
    
    async def shutdown(self):
        """Perform graceful shutdown with document cleanup."""
        if self.logger:
            self.logger.info("Starting graceful shutdown...")
        
        # Phase 1: Clean up documents and Word connections
        if self.mcp_server:
            try:
                if self.logger:
                    self.logger.info("Cleaning up documents and Word connections...")
                
                # Get Word controller and clean up documents
                word_controller = getattr(self.mcp_server, 'word_controller', None)
                if word_controller:
                    try:
                        # Save and close all open documents
                        await word_controller.cleanup_all_documents()
                        if self.logger:
                            self.logger.info("Document cleanup completed")
                    except Exception as e:
                        if self.logger:
                            self.logger.warning(f"Document cleanup failed: {e}")
                
                # Phase 2: Shutdown MCP server
                if self.logger:
                    self.logger.info("Shutting down MCP server...")
                await self.mcp_server.shutdown()
                if self.logger:
                    self.logger.info("MCP server shutdown completed")
                    
            except Exception as e:
                if self.logger:
                    self.logger.error(f"Error during MCP server shutdown: {e}")
        
        # Phase 3: Final cleanup
        if self.logger:
            self.logger.info("Graceful shutdown completed")
        
        self.shutdown_event.set()
    
    async def start_server(self, config_path: Optional[str] = None, verbose: bool = False):
        """Start the MCP server with proper initialization sequence.
        
        Args:
            config_path: Optional path to configuration file
            verbose: Enable verbose logging
        """
        startup_phase = "initialization"
        
        try:
            # Phase 1: Load and validate configuration
            startup_phase = "configuration loading"
            
            config_manager = ConfigManager(config_path)
            config = config_manager.load_config()
            
            # Override log level if verbose
            if verbose:
                config.logging.level = "DEBUG"
            
            # Phase 2: Setup logging
            startup_phase = "logging setup"
            setup_logging(config.logging)
            self.logger = logging.getLogger(__name__)
            self.logger.info("Word MCP Server starting up...")
            
            # Phase 3: Log configuration (don't print to stdout for MCP)
            self.logger.info(f"Configuration loaded from: {config_manager.get_config_path() or 'defaults'}")
            self.logger.info(f"Server config - Host: {config.server.host}, Port: {config.server.port}")
            self.logger.info(f"Max concurrent docs: {config.server.max_concurrent_docs}")
            self.logger.info(f"Word auto-launch: {config.word.auto_launch}, visible: {config.word.visible}")
            self.logger.info(f"Backup enabled: {config.word.backup_enabled}")
            self.logger.info(f"Logging level: {config.logging.level}")
            self.logger.info(f"Max file size: {config.security.max_file_size_mb}MB")
            self.logger.info(f"Allowed paths: {len(config.security.allowed_paths)} configured")
            
            # Phase 4: Setup signal handlers
            startup_phase = "signal handler setup"
            await self.setup_signal_handlers()
            self.logger.info("Signal handlers configured")
            
            # Phase 5: Initialize MCP server
            startup_phase = "MCP server initialization"
            from .server.mcp_server import WordMCPServer
            
            self.mcp_server = WordMCPServer(config_manager)
            self.logger.info("MCP server initialized")
            
            # Phase 6: Start MCP server
            startup_phase = "MCP server startup"
            self.logger.info("Starting MCP server...")
            
            await self.mcp_server.start()
            self.logger.info("MCP server started successfully")
            
        except KeyboardInterrupt:
            print("\n✓ Shutdown requested by user")
            await self.shutdown()
        except WordMCPError as e:
            error_msg = f"Word MCP Server error during {startup_phase}: {e}"
            if self.logger:
                self.logger.error(error_msg)
            print(f"✗ {error_msg}", file=sys.stderr)
            sys.exit(1)
        except Exception as e:
            error_msg = f"Unexpected error during {startup_phase}: {e}"
            if self.logger:
                self.logger.error(error_msg, exc_info=True)
            print(f"✗ {error_msg}", file=sys.stderr)
            sys.exit(1)


async def main_async(config_path: Optional[str] = None, verbose: bool = False) -> None:
    """Main async entry point.
    
    Args:
        config_path: Optional path to configuration file
        verbose: Enable verbose logging
    """
    server_manager = ServerManager()
    await server_manager.start_server(config_path, verbose)


def main() -> None:
    """Main entry point."""
    args = parse_arguments()
    
    # Handle setup and utility commands
    if args.setup:
        from .utils.setup import SetupManager
        setup_manager = SetupManager()
        setup_manager.run_full_setup()
        return
    
    if args.check_requirements:
        from .utils.setup import SetupManager
        setup_manager = SetupManager()
        requirements_met, issues = setup_manager.check_system_requirements()
        if requirements_met:
            print("✓ All system requirements met")
        else:
            print("✗ System requirements not met:")
            for issue in issues:
                print(f"  - {issue}")
            sys.exit(1)
        return
    
    if args.uninstall:
        from .utils.setup import SetupManager
        setup_manager = SetupManager()
        success = setup_manager.uninstall()
        sys.exit(0 if success else 1)
    
    if args.create_config:
        create_default_config(args.config_output)
        return
    
    try:
        asyncio.run(main_async(args.config, args.verbose))
    except KeyboardInterrupt:
        print("\n✓ Server stopped by user")
    except Exception as e:
        print(f"✗ Unexpected error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()