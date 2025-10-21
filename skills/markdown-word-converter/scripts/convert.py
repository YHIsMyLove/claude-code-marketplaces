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

        # Check for mmdc (mermaid-cli) with more reliable detection
        if not self._check_mmdc():
            missing_tools.append("mmdc")

        return missing_tools

    def _check_mmdc(self):
        """Check if mmdc is properly installed and functional"""
        try:
            result = subprocess.run(
                ["mmdc", "--version"],
                capture_output=True,
                text=True,
                timeout=5
            )
            return result.returncode == 0
        except (subprocess.TimeoutExpired, FileNotFoundError, subprocess.SubprocessError):
            return False

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

            # Build mmdc command with more compatible parameters
            cmd = [
                "mmdc",
                "-i", str(input_file),
                "-o", str(output_file),
                "-e", "png"
                # Removed "-s", "2" (scale) as it may not be supported in all versions
            ]

            print(f"Running command: {' '.join(cmd)}")

            # Run with longer timeout and better error handling
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=60  # Increased timeout for complex diagrams
            )

            if result.returncode != 0:
                print(f"✗ mmdc conversion failed:")
                print(f"  Return code: {result.returncode}")
                print(f"  Error: {result.stderr}")
                if result.stdout:
                    print(f"  Output: {result.stdout}")

                # Provide specific troubleshooting advice
                self._provide_mmdc_troubleshooting(result.stderr)
                return False
            else:
                print("✓ Mermaid diagrams converted successfully")
                if result.stdout:
                    print(f"  mmdc output: {result.stdout}")
                return True

        except subprocess.TimeoutExpired:
            print("✗ mmdc conversion timed out - diagram may be too complex")
            return False
        except FileNotFoundError:
            print("✗ mmdc command not found. Please install mermaid-cli:")
            print("  npm install -g @mermaid-js/mermaid-cli")
            return False
        except Exception as e:
            print(f"✗ Unexpected error during Mermaid conversion: {e}")
            return False

    def _provide_mmdc_troubleshooting(self, error_msg):
        """Provide specific troubleshooting advice based on error message"""
        error_lower = error_msg.lower()

        if "puppeteer" in error_lower:
            print("  💡 This appears to be a Puppeteer issue:")
            print("     - Try reinstalling mermaid-cli: npm install -g @mermaid-js/mermaid-cli")
            print("     - Or install puppeteer separately: npm install -g puppeteer")

        if "chrome" in error_lower or "chromium" in error_lower:
            print("  💡 Chrome/Chromium issue detected:")
            print("     - Ensure Chrome or Chromium is installed")
            print("     - Try: npm install -g puppeteer")

        if "syntax" in error_lower or "parse" in error_lower:
            print("  💡 Mermaid syntax issue:")
            print("     - Check your Mermaid diagram syntax")
            print("     - Ensure all code blocks are properly closed")

        if "permission" in error_lower or "access" in error_lower:
            print("  💡 Permission issue:")
            print("     - Check file permissions")
            print("     - Ensure output directory is writable")

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
            print(f"✗ Error: Input file not found: {input_file}")
            return False

        print(f"🔄 Starting conversion of: {input_path}")

        # Check dependencies with detailed feedback
        missing_tools = self.check_dependencies()
        if missing_tools:
            print(f"\n✗ Missing dependencies: {', '.join(missing_tools)}")
            print("\nTo install missing dependencies:")

            if "pandoc" in missing_tools:
                print("  • Pandoc: https://pandoc.org/installing.html")
                print("    - Windows: Download from pandoc.org")
                print("    - macOS: brew install pandoc")
                print("    - Linux: sudo apt install pandoc")

            if "mmdc" in missing_tools:
                print("  • Mermaid CLI: npm install -g @mermaid-js/mermaid-cli")
                print("    If already installed, try:")
                print("    - npm list -g @mermaid-js/mermaid-cli")
                print("    - Restart your terminal/command prompt")

            return False

        print("✓ All dependencies are available")

        # Generate output filename if not provided
        if output_file is None:
            output_file = input_path.with_suffix('.docx')

        output_path = Path(output_file)
        print(f"📄 Output will be saved to: {output_path}")

        # Create temporary directory for intermediate files
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            print(f"📁 Using temporary directory: {temp_path}")

            # Step 1: Convert Mermaid diagrams
            temp_md = temp_path / f"{input_path.stem}_processed.md"
            print("\n🎨 Step 1: Processing Mermaid diagrams...")
            mermaid_success = self.convert_mermaid_diagrams(input_path, temp_md)

            # Choose which markdown file to use for conversion
            if mermaid_success:
                md_for_conversion = temp_md
                print("✓ Using processed Markdown with converted diagrams")
            else:
                md_for_conversion = input_path
                print("! Using original Markdown (diagrams not converted)")

            # Step 2: Convert to DOCX
            print(f"\n📝 Step 2: Converting to Word document...")
            success = self.convert_to_docx(md_for_conversion, output_path, use_template)

            if success:
                print(f"\n🎉 Conversion completed successfully!")
                print(f"📄 Output file: {output_path}")
                print(f"📊 File size: {output_path.stat().st_size:,} bytes")

                if mermaid_success:
                    print("✅ Mermaid diagrams converted to PNG images")
                else:
                    print("⚠️  Mermaid diagrams were not converted (see errors above)")

                # Provide additional info
                print(f"\n💡 Next steps:")
                print(f"   • Open the document in Word to verify formatting")
                print(f"   • Check that images appear correctly")
                print(f"   • Verify table of contents is generated")

                return True
            else:
                print("\n❌ Conversion failed")
                print(f"💡 Troubleshooting:")
                print(f"   • Check that the input Markdown syntax is valid")
                print(f"   • Try converting without template: --no-template")
                print(f"   • Verify file permissions for output directory")
                return False


def test_mmdc_functionality():
    """Test mmdc functionality with a simple diagram"""
    print("🧪 Testing mmdc functionality with a simple diagram...")

    # Create a temporary Mermaid diagram file
    test_mermaid = """graph TD
    A[Start] --> B{mmdc working?}
    B -->|Yes| C[✓ Success!]
    B -->|No| D[✗ Failed]
    C --> E[End]
    D --> E
"""

    try:
        with tempfile.NamedTemporaryFile(mode='w', suffix='.mmd', delete=False) as f:
            f.write(test_mermaid)
            temp_mmd = f.name

        with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as f:
            temp_png = f.name

        print(f"📝 Test diagram created: {temp_mmd}")
        print(f"🖼️  Expected output: {temp_png}")

        # Run mmdc test
        cmd = ["mmdc", "-i", temp_mmd, "-o", temp_png, "-e", "png"]
        print(f"🔄 Running: {' '.join(cmd)}")

        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=30
        )

        # Cleanup
        os.unlink(temp_mmd)
        if os.path.exists(temp_png):
            os.unlink(temp_png)

        if result.returncode == 0:
            print("✅ mmdc functionality test PASSED")
            print("   • mmdc can successfully convert Mermaid diagrams to PNG")
            print("   • All required dependencies (Chrome/Puppeteer) are working")
            return True
        else:
            print("❌ mmdc functionality test FAILED")
            print(f"   Return code: {result.returncode}")
            print(f"   Error: {result.stderr}")
            return False

    except Exception as e:
        print(f"❌ mmdc test failed with exception: {e}")
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
    parser.add_argument(
        "--test-mmdc",
        action="store_true",
        help="Test mmdc functionality with a simple diagram"
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

    if args.test_mmdc:
        print("🔍 Testing mmdc installation and functionality...")
        print("=" * 50)

        # Basic detection
        if not shutil.which("mmdc"):
            print("❌ mmdc command not found in PATH")
            print("\n💡 Installation instructions:")
            print("   npm install -g @mermaid-js/mermaid-cli")
            print("   Then restart your terminal/command prompt")
            sys.exit(1)

        print("✅ mmdc command found in PATH")

        # Version check
        try:
            result = subprocess.run(["mmdc", "--version"], capture_output=True, text=True, timeout=10)
            if result.returncode == 0:
                print(f"✅ mmdc version: {result.stdout.strip()}")
            else:
                print(f"⚠️  mmdc version check failed: {result.stderr}")
        except Exception as e:
            print(f"⚠️  Could not get mmdc version: {e}")

        # Functionality test
        print()
        success = test_mmdc_functionality()

        if success:
            print("\n🎉 mmdc is fully functional!")
            print("   You can now convert Markdown files with Mermaid diagrams.")
        else:
            print("\n❌ mmdc has issues that need to be resolved.")
            print("\n💡 Common solutions:")
            print("   1. Reinstall mmdc: npm install -g @mermaid-js/mermaid-cli")
            print("   2. Install puppeteer: npm install -g puppeteer")
            print("   3. Ensure Chrome/Chromium is installed")
            print("   4. Try running as administrator/sudo")

        sys.exit(0 if success else 1)

    if not args.input_file:
        parser.error("input_file is required when not using --check-deps or --test-mmdc")

    success = converter.convert(
        args.input_file,
        args.output,
        use_template=not args.no_template
    )

    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()