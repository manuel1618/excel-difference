#!/usr/bin/env python3
"""
PyInstaller script for Excel Difference Generator GUI.
"""

import os
import shutil
import subprocess
import sys
from pathlib import Path


def run_pyinstaller():
    """Run PyInstaller to create executable."""

    # Get the current directory
    current_dir = Path(__file__).parent
    gui_file = current_dir / "excel_difference" / "gui.py"

    # Check if GUI file exists
    if not gui_file.exists():
        print(f"Error: GUI file not found at {gui_file}")
        return False

    # PyInstaller command
    cmd = [
        "pyinstaller",
        "--onefile",  # Create single executable
        "--windowed",  # Hide console window (Windows)
        "--name=ExcelDiffGenerator",  # Name of the executable
        "--icon=icon.ico",  # Icon file (if exists)
        "--add-data=data;data",  # Include data directory
        "--hidden-import=openpyxl",  # Include openpyxl
        "--hidden-import=pandas",  # Include pandas
        "--hidden-import=tkinter",  # Include tkinter
        "--hidden-import=tkinter.ttk",  # Include ttk
        "--hidden-import=tkinter.filedialog",  # Include filedialog
        "--hidden-import=tkinter.scrolledtext",  # Include scrolledtext
        "--hidden-import=tkinter.messagebox",  # Include messagebox
        str(gui_file),
    ]

    # Remove icon if it doesn't exist
    if not (current_dir / "icon.ico").exists():
        cmd = [arg for arg in cmd if not arg.startswith("--icon")]

    # Remove data directory if it doesn't exist
    if not (current_dir / "data").exists():
        cmd = [arg for arg in cmd if not arg.startswith("--add-data")]

    print("Running PyInstaller with command:")
    print(" ".join(cmd))
    print("-" * 50)

    try:
        # Run PyInstaller
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        print("PyInstaller output:")
        print(result.stdout)

        if result.stderr:
            print("PyInstaller warnings/errors:")
            print(result.stderr)

        # Check if executable was created
        exe_path = current_dir / "dist" / "ExcelDiffGenerator.exe"
        if exe_path.exists():
            print(f"\n✅ Success! Executable created at: {exe_path}")
            print(f"File size: {exe_path.stat().st_size / (1024*1024):.1f} MB")
            return True
        else:
            print(f"\n❌ Error: Executable not found at {exe_path}")
            return False

    except subprocess.CalledProcessError as e:
        print(f"❌ PyInstaller failed with error code {e.returncode}")
        print("Error output:")
        print(e.stderr)
        return False
    except FileNotFoundError:
        print("❌ PyInstaller not found. Please install it first:")
        print("pip install pyinstaller")
        return False


def clean_build():
    """Clean build artifacts."""
    current_dir = Path(__file__).parent

    # Directories to remove
    dirs_to_remove = ["build", "dist", "__pycache__"]

    for dir_name in dirs_to_remove:
        dir_path = current_dir / dir_name
        if dir_path.exists():
            print(f"Removing {dir_path}...")
            shutil.rmtree(dir_path)

    # Remove spec file
    spec_file = current_dir / "ExcelDiffGenerator.spec"
    if spec_file.exists():
        print(f"Removing {spec_file}...")
        spec_file.unlink()


def main():
    """Main function."""
    print("Excel Difference Generator - PyInstaller Build Script")
    print("=" * 50)

    # Check if we should clean first
    if len(sys.argv) > 1 and sys.argv[1] == "clean":
        print("Cleaning build artifacts...")
        clean_build()
        return

    # Check if PyInstaller is installed
    try:
        import PyInstaller

        print("PyInstaller is installed")
    except ImportError:
        print("❌ PyInstaller not installed. Installing...")
        subprocess.run(
            [sys.executable, "-m", "pip", "install", "pyinstaller"], check=True
        )
        print("✅ PyInstaller installed successfully!")

    # Run PyInstaller
    success = run_pyinstaller()

    if success:
        print("\n🎉 Build completed successfully!")
        print("\nYou can now run the executable from the 'dist' folder.")
        print(
            "The executable is self-contained and can be distributed to other machines."
        )
    else:
        print("\n💥 Build failed!")
        sys.exit(1)


if __name__ == "__main__":
    main()
