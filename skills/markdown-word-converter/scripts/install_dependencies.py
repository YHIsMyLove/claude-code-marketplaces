#!/usr/bin/env python3
"""
Dependency Installation Helper for Markdown to Word Converter
Checks and guides installation of required dependencies.
"""

import subprocess
import sys
import platform
import shutil
from pathlib import Path


class DependencyInstaller:
    def __init__(self):
        self.system = platform.system().lower()
        self.required_tools = {
            "pandoc": {
                "check_cmd": "pandoc --version",
                "install_windows": "Download from https://pandoc.org/installing.html",
                "install_macos": "brew install pandoc",
                "install_linux": {
                    "ubuntu": "sudo apt update && sudo apt install pandoc",
                    "debian": "sudo apt update && sudo apt install pandoc",
                    "fedora": "sudo dnf install pandoc",
                    "centos": "sudo yum install pandoc",
                    "arch": "sudo pacman -S pandoc"
                }
            },
            "mmdc": {
                "check_cmd": "mmdc --version",
                "install_windows": "npm install -g @mermaid-js/mermaid-cli",
                "install_macos": "npm install -g @mermaid-js/mermaid-cli",
                "install_linux": "npm install -g @mermaid-js/mermaid-cli"
            }
        }

    def check_tool(self, tool_name):
        """Check if a tool is installed"""
        tool_info = self.required_tools.get(tool_name)
        if not tool_info:
            return False

        try:
            result = subprocess.run(
                tool_info["check_cmd"].split(),
                capture_output=True,
                text=True
            )
            return result.returncode == 0
        except (FileNotFoundError, subprocess.SubprocessError):
            return False

    def check_all_dependencies(self):
        """Check all required dependencies"""
        print("Checking dependencies...")
        print("-" * 40)

        missing_tools = []
        installed_tools = []

        for tool in self.required_tools:
            if self.check_tool(tool):
                print(f"✓ {tool}: Installed")
                installed_tools.append(tool)
            else:
                print(f"✗ {tool}: Not found")
                missing_tools.append(tool)

        print("-" * 40)

        if not missing_tools:
            print("✓ All dependencies are installed!")
            return True

        print(f"Missing dependencies: {', '.join(missing_tools)}")
        return False

    def get_linux_distro(self):
        """Try to detect Linux distribution"""
        try:
            with open("/etc/os-release", "r") as f:
                for line in f:
                    if line.startswith("ID="):
                        return line.split("=")[1].strip().lower()
        except (FileNotFoundError, PermissionError):
            pass

        return "unknown"

    def show_installation_instructions(self, tool=None):
        """Show installation instructions for missing tools"""
        if tool and tool in self.required_tools:
            tools_to_show = [tool]
        else:
            tools_to_show = [
                t for t in self.required_tools
                if not self.check_tool(t)
            ]

        if not tools_to_show:
            print("All dependencies are already installed.")
            return

        print("\nInstallation Instructions:")
        print("=" * 50)

        for tool in tools_to_show:
            tool_info = self.required_tools[tool]
            print(f"\n{tool.upper()}:")
            print("-" * 20)

            if self.system == "windows":
                print(f"Windows: {tool_info['install_windows']}")
            elif self.system == "darwin":  # macOS
                print(f"macOS: {tool_info['install_macos']}")
            elif self.system == "linux":
                distro = self.get_linux_distro()
                if distro in tool_info["install_linux"]:
                    print(f"Linux ({distro}): {tool_info['install_linux'][distro]}")
                else:
                    print(f"Linux (general): {tool_info['install_linux'].get('ubuntu', tool_info['install_linux'].get('ubuntu', 'npm install -g @mermaid-js/mermaid-cli'))}")

        print("\n" + "=" * 50)
        print("After installing dependencies, run this check again to verify.")

    def attempt_auto_install(self, tool):
        """Attempt to automatically install a tool (limited functionality)"""
        if tool == "pandoc":
            print("Pandoc cannot be auto-installed by this script.")
            print("Please follow the manual installation instructions above.")
            return False

        elif tool == "mmdc":
            if not shutil.which("npm"):
                print("npm is required to install mmdc. Please install Node.js first.")
                return False

            try:
                print(f"Attempting to install {tool}...")
                cmd = ["npm", "install", "-g", "@mermaid-js/mermaid-cli"]
                result = subprocess.run(cmd, capture_output=True, text=True)

                if result.returncode == 0:
                    print(f"✓ {tool} installed successfully!")
                    return True
                else:
                    print(f"Failed to install {tool}: {result.stderr}")
                    return False

            except subprocess.SubprocessError as e:
                print(f"Error during installation: {e}")
                return False

        return False


def main():
    installer = DependencyInstaller()

    if len(sys.argv) > 1:
        if sys.argv[1] == "--install":
            if len(sys.argv) > 2:
                tool = sys.argv[2]
                if tool in installer.required_tools:
                    installer.attempt_auto_install(tool)
                else:
                    print(f"Unknown tool: {tool}")
            else:
                print("Please specify a tool to install: --install <tool>")
        elif sys.argv[1] == "--instructions":
            if len(sys.argv) > 2:
                installer.show_installation_instructions(sys.argv[2])
            else:
                installer.show_installation_instructions()
        else:
            print("Usage:")
            print("  python install_dependencies.py [--install <tool>] [--instructions [<tool>]]")
    else:
        # Default behavior: check all dependencies
        all_installed = installer.check_all_dependencies()

        if not all_installed:
            installer.show_installation_instructions()
            sys.exit(1)
        else:
            sys.exit(0)


if __name__ == "__main__":
    main()