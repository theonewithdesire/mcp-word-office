#!/usr/bin/env python3
"""
Advanced Document Features Example

This example demonstrates advanced features of the Word MCP server:
- Creating tables with formatting
- Working with headers and footers
- Page formatting and breaks
- Find and replace operations
- Document reading and analysis

Usage:
    python examples/advanced_document_features.py
"""

import asyncio
import sys
import os
from pathlib import Path

# Add the project root to the Python path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from word_mcp_server.server.mcp_server import WordMCPServer
from word_mcp_server.config.config_manager import ConfigManager


async def create_report_document():
    """Create a comprehensive report document with advanced features."""
    print("=== Creating Advanced Report Document ===\n")
    
    try:
        # Initialize server
        config_manager = ConfigManager()
        server = WordMCPServer(config_manager)
        
        # Create document
        print("1. Creating new report document...")
        doc_id = await server._handle_create_document(title="Quarterly Sales Report")
        print(f"   ✓ Document created: {doc_id}\n")
        
        # Set up headers and footers
        print("2. Setting up headers and footers...")
        await server._handle_insert_header_footer(
            doc_id=doc_id,
            header_text="Quarterly Sales Report - Q4 2024",
            footer_text="Confidential - Internal Use Only | Page ",
            section_index=1
        )
        print("   ✓ Headers and footers configured\n")
        
        # Add document title
        print("3. Adding document title...")
        title_text = "QUARTERLY SALES REPORT\nQ4 2024\n\n"
        await server._handle_insert_text(
            doc_id=doc_id,
            text=title_text,
            position=0
        )
        
        # Format title
        await server._handle_format_text(
            doc_id=doc_id,
            start=0,
            end=len("QUARTERLY SALES REPORT"),
            bold=True,
            font_size=18,
            color="blue"
        )
        print("   ✓ Title added and formatted\n")
        
        # Add executive summary
        print("4. Adding executive summary...")
        summary_text = """EXECUTIVE SUMMARY
        
This report presents the sales performance for Q4 2024, highlighting key achievements, challenges, and strategic recommendations for the upcoming quarter.

Key Highlights:
• Total revenue increased by 15% compared to Q3 2024
• Customer acquisition grew by 22%
• Product line expansion contributed to 8% revenue growth
• Regional performance varied significantly

"""
        
        current_pos = len(title_text)
        await server._handle_insert_text(
            doc_id=doc_id,
            text=summary_text,
            position=current_pos
        )
        
        # Format executive summary header
        summary_start = current_pos
        summary_header_end = summary_start + len("EXECUTIVE SUMMARY")
        await server._handle_format_text(
            doc_id=doc_id,
            start=summary_start,
            end=summary_header_end,
            bold=True,
            font_size=14,
            underline=True
        )
        print("   ✓ Executive summary added\n")
        
        # Insert page break
        print("5. Inserting page break...")
        current_pos += len(summary_text)
        await server._handle_insert_page_break(
            doc_id=doc_id,
            position=current_pos,
            break_type="page"
        )
        print("   ✓ Page break inserted\n")
        
        # Add sales data table
        print("6. Creating sales data table...")
        table_header_text = "\n\nSALES PERFORMANCE BY REGION\n\n"
        await server._handle_insert_text(
            doc_id=doc_id,
            text=table_header_text,
            position=current_pos
        )
        
        current_pos += len(table_header_text)
        
        # Create table
        table_result = await server._handle_create_table(
            doc_id=doc_id,
            rows=5,
            cols=4,
            position=current_pos
        )
        print(f"   ✓ Table created: {table_result}\n")
        
        # Add recommendations section
        print("7. Adding recommendations section...")
        recommendations_text = """

STRATEGIC RECOMMENDATIONS

Based on the Q4 2024 performance analysis, we recommend the following strategic initiatives:

1. MARKET EXPANSION
   - Focus on underperforming regions
   - Increase marketing investment in high-potential areas
   - Develop region-specific product offerings

2. PRODUCT DEVELOPMENT
   - Accelerate innovation in core product lines
   - Invest in emerging technology solutions
   - Enhance customer feedback integration

3. OPERATIONAL EFFICIENCY
   - Streamline sales processes
   - Implement advanced analytics tools
   - Optimize resource allocation

"""
        
        await server._handle_insert_text(
            doc_id=doc_id,
            text=recommendations_text,
            position=current_pos + 200  # Approximate position after table
        )
        print("   ✓ Recommendations section added\n")
        
        # Perform find and replace operation
        print("8. Performing find and replace...")
        replace_result = await server._handle_find_replace(
            doc_id=doc_id,
            find_text="Q4 2024",
            replace_text="Fourth Quarter 2024",
            match_case=False
        )
        print(f"   ✓ Find and replace completed: {replace_result}\n")
        
        # Set page formatting
        print("9. Setting page formatting...")
        await server._handle_set_page_formatting(
            doc_id=doc_id,
            section_index=1,
            margins={"top": 72, "bottom": 72, "left": 90, "right": 90},
            orientation="portrait",
            paper_size="letter"
        )
        print("   ✓ Page formatting applied\n")
        
        # Save document
        print("10. Saving the report...")
        save_path = project_root / "examples" / "output" / "quarterly_sales_report.docx"
        save_path.parent.mkdir(exist_ok=True)
        
        await server._handle_save_document(
            doc_id=doc_id,
            path=str(save_path)
        )
        print(f"    ✓ Report saved to: {save_path}\n")
        
        # Close document
        await server._handle_close_document(doc_id=doc_id, save=True)
        print("    ✓ Document closed\n")
        
        return str(save_path)
        
    except Exception as e:
        print(f"❌ Error creating report: {e}")
        import traceback
        traceback.print_exc()
        return None


async def analyze_document():
    """Demonstrate document reading and analysis capabilities."""
    print("=== Document Analysis Example ===\n")
    
    try:
        config_manager = ConfigManager()
        server = WordMCPServer(config_manager)
        
        # Create a sample document to analyze
        print("1. Creating sample document for analysis...")
        doc_path = project_root / "examples" / "output" / "sample_analysis.docx"
        doc_path.parent.mkdir(exist_ok=True)
        
        # For this example, we'll create a simple document
        doc_id = await server._handle_create_document(title="Sample Document")
        
        sample_content = """Sample Document for Analysis

This is a sample document that contains various elements for analysis:

1. Multiple paragraphs
2. Different formatting styles
3. Lists and structured content

The document demonstrates the capabilities of the Word MCP server in reading and analyzing document content.

Key features include:
• Text extraction
• Structure analysis
• Metadata retrieval
• Statistical information

This content will be analyzed to demonstrate the document reading capabilities."""
        
        await server._handle_insert_text(
            doc_id=doc_id,
            text=sample_content,
            position=0
        )
        
        await server._handle_save_document(doc_id=doc_id, path=str(doc_path))
        await server._handle_close_document(doc_id=doc_id, save=True)
        print(f"   ✓ Sample document created: {doc_path}\n")
        
        # Now analyze the document
        print("2. Reading document content...")
        read_result = await server._handle_read_document(path=str(doc_path))
        print(f"   ✓ Document content read: {len(read_result.get('text', ''))} characters\n")
        
        # Get document information
        print("3. Retrieving document metadata...")
        info_result = await server._handle_get_document_info(path=str(doc_path))
        print(f"   ✓ Document metadata retrieved\n")
        
        # Get document statistics
        print("4. Calculating document statistics...")
        stats_result = await server._handle_get_document_statistics(path=str(doc_path))
        print(f"   ✓ Document statistics calculated\n")
        
        # Display analysis results
        print("=== Analysis Results ===")
        if isinstance(read_result, dict):
            print(f"Text length: {len(read_result.get('text', ''))}")
            print(f"Paragraphs: {len(read_result.get('paragraphs', []))}")
        
        if isinstance(info_result, dict):
            metadata = info_result.get('metadata', {})
            print(f"Title: {metadata.get('title', 'N/A')}")
            print(f"Author: {metadata.get('author', 'N/A')}")
        
        if isinstance(stats_result, dict):
            print(f"Word count: {stats_result.get('word_count', 'N/A')}")
            print(f"Character count: {stats_result.get('character_count', 'N/A')}")
        
        print("\n=== Document analysis completed ===")
        
    except Exception as e:
        print(f"❌ Error analyzing document: {e}")
        import traceback
        traceback.print_exc()


async def demonstrate_concurrent_operations():
    """Demonstrate handling multiple documents concurrently."""
    print("\n=== Concurrent Operations Example ===\n")
    
    try:
        config_manager = ConfigManager()
        server = WordMCPServer(config_manager)
        
        print("1. Creating multiple documents concurrently...")
        
        # Create multiple documents at once
        create_tasks = [
            server._handle_create_document(title=f"Concurrent Document {i}")
            for i in range(3)
        ]
        
        doc_ids = await asyncio.gather(*create_tasks)
        print(f"   ✓ Created {len(doc_ids)} documents concurrently\n")
        
        # Add content to all documents concurrently
        print("2. Adding content to all documents...")
        content_tasks = []
        for i, doc_id in enumerate(doc_ids):
            content_tasks.append(
                server._handle_insert_text(
                    doc_id=doc_id,
                    text=f"This is content for concurrent document {i+1}.\n\nThis demonstrates the ability to handle multiple document operations simultaneously.",
                    position=0
                )
            )
        
        await asyncio.gather(*content_tasks)
        print("   ✓ Content added to all documents\n")
        
        # Save all documents concurrently
        print("3. Saving all documents...")
        save_tasks = []
        for i, doc_id in enumerate(doc_ids):
            save_path = project_root / "examples" / "output" / f"concurrent_doc_{i+1}.docx"
            save_tasks.append(
                server._handle_save_document(doc_id=doc_id, path=str(save_path))
            )
        
        await asyncio.gather(*save_tasks)
        print("   ✓ All documents saved\n")
        
        # Close all documents
        print("4. Closing all documents...")
        close_tasks = [
            server._handle_close_document(doc_id=doc_id, save=True)
            for doc_id in doc_ids
        ]
        
        await asyncio.gather(*close_tasks)
        print("   ✓ All documents closed\n")
        
        print("=== Concurrent operations completed successfully ===")
        
    except Exception as e:
        print(f"❌ Error in concurrent operations: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    print("Starting Word MCP Server Advanced Features Example...")
    print("This example demonstrates advanced document manipulation capabilities.\n")
    
    async def run_all_examples():
        # Create advanced report
        report_path = await create_report_document()
        
        # Analyze document
        await analyze_document()
        
        # Demonstrate concurrent operations
        await demonstrate_concurrent_operations()
        
        print("\n=== All Advanced Examples Completed ===")
        if report_path:
            print(f"Main report saved to: {report_path}")
        print("Check the examples/output/ directory for all generated files.")
    
    # Run all examples
    asyncio.run(run_all_examples())