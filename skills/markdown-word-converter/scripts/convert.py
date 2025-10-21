#!/usr/bin/env python3
"""
Markdown to Word Converter
A cross-platform Python script to convert Markdown files to Word documents
with support for Mermaid diagrams and custom templates.
"""

import argparse
import os
import sys
import subprocess
import shutil
from pathlib import Path
import tempfile


class MarkdownToWordConverter:
    def __init__(self):
        self.script_dir = Path(__file__).parent
        self.assets_dir = self.script_dir.parent / "assets"
        self.template_docx = self.assets_dir / "template.docx"

    def check_dependencies(self):
        """Check if required tools are installed"""
        missing_tools = []

        # Check for pandoc
        if not shutil.which("pandoc"):
            missing_tools.append("pandoc")

        # Check for mmdc (mermaid-cli)
        if not shutil.which("mmdc"):
            missing_tools.append("mmdc")

        return missing_tools

    def install_dependencies(self):
        """Install missing dependencies"""
        missing_tools = self.check_dependencies()

        if not missing_tools:
            print("✓ All dependencies are already installed")
            return True

        print(f"Missing tools: {', '.join(missing_tools)}")
        print("\nTo install missing tools:")

        if "pandoc" in missing_tools:
            print("• Pandoc: https://pandoc.org/installing.html")
            print("  - Windows: Download installer from pandoc.org")
            print("  - macOS: brew install pandoc")
            print("  - Linux: sudo apt install pandoc (Ubuntu/Debian)")

        if "mmdc" in missing_tools:
            print("• Mermaid CLI: https://github.com/mermaid-js/mermaid-cli")
            print("  - npm install -g @mermaid-js/mermaid-cli")

        return False

    def convert_mermaid_diagrams(self, input_file, output_file):
        """Convert Mermaid diagrams using mmdc"""
        try:
            print(f"Converting Mermaid diagrams in {input_file}...")
            cmd = [
                "mmdc",
                "-i", str(input_file),
                "-o", str(output_file),
                "-e", "png",
                "-s", "2"
            ]

            result = subprocess.run(cmd, capture_output=True, text=True)
            if result.returncode != 0:
                print(f"Warning: mmdc failed: {result.stderr}")
                return False

            return True

        except FileNotFoundError:
            print("Error: mmdc not found. Please install mermaid-cli")
            return False

    def convert_to_docx(self, input_file, output_file, use_template=True):
        """Convert markdown to docx using pandoc"""
        try:
            print(f"Converting to DOCX: {output_file}...")

            cmd = ["pandoc", str(input_file), "-o", str(output_file)]

            # Add table of contents
            cmd.append("--toc")

            # Add template if available and requested
            if use_template and self.template_docx.exists():
                cmd.extend(["--reference-doc", str(self.template_docx)])
                print(f"Using template: {self.template_docx}")
            elif use_template:
                print("Warning: Template file not found, using default styling")

            result = subprocess.run(cmd, capture_output=True, text=True)
            if result.returncode != 0:
                print(f"Error: pandoc failed: {result.stderr}")
                return False

            return True

        except FileNotFoundError:
            print("Error: pandoc not found. Please install pandoc")
            return False

    def convert(self, input_file, output_file=None, use_template=True):
        """Main conversion function"""
        input_path = Path(input_file)

        if not input_path.exists():
            print(f"Error: Input file not found: {input_file}")
            return False

        # Check dependencies
        if not self.install_dependencies():
            print("\nPlease install missing dependencies and try again.")
            return False

        # Generate output filename if not provided
        if output_file is None:
            output_file = input_path.with_suffix('.docx')

        output_path = Path(output_file)

        # Create temporary directory for intermediate files
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)

            # Step 1: Convert Mermaid diagrams
            temp_md = temp_path / f"{input_path.stem}_processed.md"
            mermaid_success = self.convert_mermaid_diagrams(input_path, temp_md)

            # Choose which markdown file to use for conversion
            md_for_conversion = temp_md if mermaid_success else input_path

            # Step 2: Convert to DOCX
            success = self.convert_to_docx(md_for_conversion, output_path, use_template)

            if success:
                print(f"\n✓ Conversion completed successfully!")
                print(f"Output file: {output_path}")

                if mermaid_success:
                    print("✓ Mermaid diagrams converted")
                else:
                    print("! Mermaid diagrams were not converted (mmdc not available)")

                return True
            else:
                print("✗ Conversion failed")
                return False


def main():
    parser = argparse.ArgumentParser(
        description="Convert Markdown files to Word documents with Mermaid diagram support"
    )
    parser.add_argument(
        "input_file",
        nargs="?",
        help="Input Markdown file path"
    )
    parser.add_argument(
        "-o", "--output",
        help="Output DOCX file path (default: same name as input with .docx extension)"
    )
    parser.add_argument(
        "--no-template",
        action="store_true",
        help="Don't use the default Word template"
    )
    parser.add_argument(
        "--check-deps",
        action="store_true",
        help="Check dependencies and exit"
    )

    args = parser.parse_args()

    converter = MarkdownToWordConverter()

    if args.check_deps:
        missing = converter.check_dependencies()
        if missing:
            print(f"Missing dependencies: {', '.join(missing)}")
            sys.exit(1)
        else:
            print("All dependencies are installed")
            sys.exit(0)

    if not args.input_file:
        parser.error("input_file is required when not using --check-deps")

    success = converter.convert(
        args.input_file,
        args.output,
        use_template=not args.no_template
    )

    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()