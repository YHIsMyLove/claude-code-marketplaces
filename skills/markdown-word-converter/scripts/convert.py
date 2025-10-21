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
import platform
from pathlib import Path
import tempfile


class MarkdownToWordConverter:
    def __init__(self, mmdc_path=None):
        self.script_dir = Path(__file__).parent
        self.assets_dir = self.script_dir.parent / "assets"
        self.template_docx = self.assets_dir / "template.docx"
        self.mmdc_path = mmdc_path  # Custom mmdc path if provided

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

    def _get_npm_global_paths(self):
        """Get npm global installation paths"""
        try:
            # Get npm global prefix
            result = subprocess.run(
                ["npm", "config", "get", "prefix"],
                capture_output=True,
                text=True,
                timeout=10
            )
            if result.returncode == 0:
                npm_prefix = result.stdout.strip()
                paths = []

                # Add platform-specific paths
                if platform.system() == "Windows":
                    paths.append(Path(npm_prefix) / "node_modules" / ".bin")
                    # Also check AppData path
                    appdata_npm = Path(os.environ.get("APPDATA", "")) / "npm"
                    if appdata_npm.exists():
                        paths.append(appdata_npm)
                else:
                    paths.append(Path(npm_prefix) / "bin")
                    # Also check common Unix paths
                    home_npm = Path.home() / ".npm" / "global" / "bin"
                    if home_npm.exists():
                        paths.append(home_npm)

                return [str(p) for p in paths if p.exists()]
        except Exception:
            pass
        return []

    def _find_mmdc_in_npm_paths(self):
        """Try to find mmdc in npm global installation paths"""
        npm_paths = self._get_npm_global_paths()

        for npm_path in npm_paths:
            mmdc_path = Path(npm_path) / ("mmdc.cmd" if platform.system() == "Windows" else "mmdc")
            if mmdc_path.exists() and mmdc_path.is_file():
                return str(mmdc_path)

        # Also check for mmdc in node_modules/.bin relative to mermaid-cli installation
        try:
            result = subprocess.run(
                ["npm", "list", "-g", "@mermaid-js/mermaid-cli"],
                capture_output=True,
                text=True,
                timeout=10
            )
            if result.returncode == 0:
                # Parse npm output to find installation path
                lines = result.stdout.strip().split('\n')
                for line in lines:
                    if "@mermaid-js/mermaid-cli@" in line:
                        # Extract path if available
                        parts = line.split()
                        for part in parts:
                            if part.startswith("node_modules") or os.path.isabs(part):
                                if os.path.isabs(part):
                                    mermaid_path = Path(part)
                                    bin_path = mermaid_path / "node_modules" / ".bin"
                                    if bin_path.exists():
                                        mmdc_path = bin_path / ("mmdc.cmd" if platform.system() == "Windows" else "mmdc")
                                        if mmdc_path.exists():
                                            return str(mmdc_path)
        except Exception:
            pass

        return None

    def _check_mmdc(self):
        """Check if mmdc is properly installed and functional with enhanced detection"""
        # Method 1: Use custom path if provided
        if self.mmdc_path:
            if os.path.isfile(self.mmdc_path):
                return self._test_mmdc_executable(self.mmdc_path)
            else:
                print(f"Warning: Custom mmdc path not found: {self.mmdc_path}")

        # Method 2: Standard PATH check
        if shutil.which("mmdc"):
            return self._test_mmdc_executable("mmdc")

        # Method 3: Enhanced npm path detection
        mmdc_path = self._find_mmdc_in_npm_paths()
        if mmdc_path:
            print(f"Found mmdc at: {mmdc_path}")
            if self._test_mmdc_executable(mmdc_path):
                # Store the found path for later use
                self.mmdc_path = mmdc_path
                return True

        # Method 4: Try common Windows paths directly (Windows specific)
        if platform.system() == "Windows":
            common_paths = [
                Path(os.environ.get("APPDATA", "")) / "npm" / "mmdc.cmd",
                Path("C:") / "Users" / os.environ.get("USERNAME", "") / "AppData" / "Roaming" / "npm" / "mmdc.cmd",
                Path("C:") / "Program Files" / "nodejs" / "node_modules" / ".bin" / "mmdc.cmd",
                Path("C:") / "Program Files (x86)" / "nodejs" / "node_modules" / ".bin" / "mmdc.cmd"
            ]

            for path in common_paths:
                if path.exists() and path.is_file():
                    print(f"Found mmdc at common path: {path}")
                    if self._test_mmdc_executable(str(path)):
                        self.mmdc_path = str(path)
                        return True

        # Method 5: Try npm to run mmdc directly (last resort)
        try:
            result = subprocess.run(
                ["npm", "run", "--silent", "mmdc", "--", "--version"],
                capture_output=True,
                text=True,
                timeout=15
            )
            if result.returncode == 0:
                print("mmdc accessible via npm run")
                self.mmdc_path = "npm run --silent mmdc --"
                return True
        except Exception:
            pass

        return False

    def _test_mmdc_executable(self, mmdc_cmd):
        """Test if mmdc executable works properly"""
        try:
            # Handle npm run command format
            if isinstance(mmdc_cmd, str) and "npm run" in mmdc_cmd:
                cmd = mmdc_cmd.split() + ["--version"]
            else:
                cmd = [mmdc_cmd, "--version"]

            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=30  # Increased timeout for first run
            )

            if result.returncode == 0:
                print(f"✅ mmdc version: {result.stdout.strip()}")
                return True
            else:
                print(f"⚠️  mmdc test failed: {result.stderr}")
                return False

        except subprocess.TimeoutExpired:
            print("⚠️  mmdc version check timed out (may be downloading dependencies)")
            return False
        except (FileNotFoundError, subprocess.SubprocessError) as e:
            print(f"⚠️  mmdc test failed: {e}")
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
        """Convert Mermaid diagrams using mmdc with enhanced path support"""
        try:
            print(f"Converting Mermaid diagrams in {input_file}...")

            # Build mmdc command based on detected path
            if self.mmdc_path and "npm run" in self.mmdc_path:
                # Use npm run command format
                cmd = self.mmdc_path.split() + ["-i", str(input_file), "-o", str(output_file), "-e", "png"]
            elif self.mmdc_path and os.path.isfile(self.mmdc_path):
                # Use custom detected path
                cmd = [self.mmdc_path, "-i", str(input_file), "-o", str(output_file), "-e", "png"]
            else:
                # Use standard command (fallback)
                cmd = ["mmdc", "-i", str(input_file), "-o", str(output_file), "-e", "png"]

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

    def _provide_environment_troubleshooting(self):
        """Provide environment-specific troubleshooting for mmdc detection issues"""
        print("\n  🔍 Environment troubleshooting:")

        # Check npm installation
        try:
            result = subprocess.run(["npm", "--version"], capture_output=True, text=True, timeout=5)
            if result.returncode == 0:
                print(f"     ✅ npm version: {result.stdout.strip()}")
            else:
                print("     ❌ npm not working properly")
        except Exception:
            print("     ❌ npm not found or not working")

        # Check if mermaid-cli is installed via npm
        try:
            result = subprocess.run(
                ["npm", "list", "-g", "@mermaid-js/mermaid-cli"],
                capture_output=True,
                text=True,
                timeout=10
            )
            if result.returncode == 0:
                print("     ✅ @mermaid-js/mermaid-cli is installed globally")
                lines = result.stdout.strip().split('\n')
                for line in lines:
                    if "@mermaid-js/mermaid-cli@" in line and "empty" not in line.lower():
                        print(f"     📦 Found: {line.strip()}")
            else:
                print("     ❌ @mermaid-js/mermaid-cli not found in global packages")
        except Exception as e:
            print(f"     ⚠️  Could not check npm packages: {e}")

        # Show npm global paths
        npm_paths = self._get_npm_global_paths()
        if npm_paths:
            print("     📂 npm global paths:")
            for path in npm_paths:
                print(f"        - {path}")
                # Check if mmdc exists in each path
                mmdc_in_path = Path(path) / ("mmdc.cmd" if platform.system() == "Windows" else "mmdc")
                if mmdc_in_path.exists():
                    print(f"          ✅ mmdc found here!")
                else:
                    print(f"          ❌ mmdc not found here")
        else:
            print("     ⚠️  Could not determine npm global paths")

        # Show PATH environment variable
        print("     🛤️  Current PATH includes:")
        path_dirs = os.environ.get("PATH", "").split(os.pathsep)
        npm_in_path = any("npm" in path_dir.lower() for path_dir in path_dirs)
        if npm_in_path:
            for path_dir in path_dirs:
                if "npm" in path_dir.lower():
                    print(f"        - {path_dir}")
        else:
            print("        ❌ No npm-related directories found in PATH")

        # Platform-specific advice
        if platform.system() == "Windows":
            print("\n     🪟 Windows-specific advice:")
            print("        • Check if npm global path is in system PATH")
            print("        • Try restarting Command Prompt as Administrator")
            print("        • Verify npm installation: npm config get prefix")
            print("        • Manual fix: Add npm path to system PATH")
        else:
            print("\n     🐧 Unix-specific advice:")
            print("        • Check shell profile: ~/.bashrc, ~/.zshrc, etc.")
            print("        • Try: export PATH=\"$(npm config get prefix)/bin:$PATH\"")
            print("        • Restart terminal or run: source ~/.bashrc")

        print("\n     💡 Quick fix options:")
        print("        1. Use --mmdc-path to specify mmdc location directly")
        print("        2. Add npm global path to PATH environment variable")
        print("        3. Reinstall mermaid-cli: npm install -g @mermaid-js/mermaid-cli")

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

                # Provide environment-specific troubleshooting
                self._provide_environment_troubleshooting()

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


def test_mmdc_functionality(mmdc_path=None):
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

        # Build mmdc command based on provided path
        if mmdc_path and "npm run" in mmdc_path:
            cmd = mmdc_path.split() + ["-i", temp_mmd, "-o", temp_png, "-e", "png"]
        elif mmdc_path and os.path.isfile(mmdc_path):
            cmd = [mmdc_path, "-i", temp_mmd, "-o", temp_png, "-e", "png"]
        else:
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
    parser.add_argument(
        "--mmdc-path",
        help="Specify path to mmdc executable (useful if not in PATH)"
    )

    args = parser.parse_args()

    converter = MarkdownToWordConverter(mmdc_path=args.mmdc_path)

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

        # Use enhanced detection
        mmdc_found = converter._check_mmdc()

        if not mmdc_found:
            print("\n❌ mmdc detection failed")
            print("\n💡 Installation instructions:")
            print("   npm install -g @mermaid-js/mermaid-cli")
            print("   Then restart your terminal/command prompt")

            # Provide troubleshooting
            converter._provide_environment_troubleshooting()
            sys.exit(1)

        print("✅ mmdc detected successfully")

        # Functionality test
        print()
        success = test_mmdc_functionality(mmdc_path=converter.mmdc_path)

        if success:
            print("\n🎉 mmdc is fully functional!")
            print("   You can now convert Markdown files with Mermaid diagrams.")
            if converter.mmdc_path:
                print(f"   Using mmdc from: {converter.mmdc_path}")
        else:
            print("\n❌ mmdc has issues that need to be resolved.")
            print("\n💡 Common solutions:")
            print("   1. Reinstall mmdc: npm install -g @mermaid-js/mermaid-cli")
            print("   2. Install puppeteer: npm install -g puppeteer")
            print("   3. Ensure Chrome/Chromium is installed")
            print("   4. Try running as administrator/sudo")
            print("   5. Use --mmdc-path to specify correct path")

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