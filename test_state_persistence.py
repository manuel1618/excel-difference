#!/usr/bin/env python3
"""
Test script to verify state persistence functionality.
"""

import json
import tempfile
from pathlib import Path
from unittest.mock import patch

from excel_difference.gui import StateManager


def test_state_manager():
    """Test the StateManager class functionality."""

    # Create a temporary directory for testing
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)

        # Create some test files
        test_file1 = temp_path / "file1.xlsx"
        test_file2 = temp_path / "file2.xlsx"
        test_file1.touch()
        test_file2.touch()

        # Mock the home directory to use our temp directory
        with patch.object(Path, "home", return_value=temp_path):
            state_manager = StateManager()

            # Test saving state
            test_config = {
                "file1_path": str(test_file1),
                "file2_path": str(test_file2),
                "output_path": str(temp_path / "output.xlsx"),
                "key_column": 2,
                "key_row": 3,
            }

            state_manager.save_state(
                test_config["file1_path"],
                test_config["file2_path"],
                test_config["output_path"],
                test_config["key_column"],
                test_config["key_row"],
            )

            # Verify config file was created
            assert state_manager.config_file.exists()

            # Test loading state
            loaded_config = state_manager.load_state()

            # Verify loaded config matches saved config
            assert loaded_config["file1_path"] == test_config["file1_path"]
            assert loaded_config["file2_path"] == test_config["file2_path"]
            assert loaded_config["output_path"] == test_config["output_path"]
            assert loaded_config["key_column"] == test_config["key_column"]
            assert loaded_config["key_row"] == test_config["key_row"]

            print("✅ State persistence test passed!")

            # Test file validation - save with non-existent files
            state_manager.save_state(
                "/non/existent/file1.xlsx",
                "/non/existent/file2.xlsx",
                test_config["output_path"],
                test_config["key_column"],
                test_config["key_row"],
            )

            # Load should clear non-existent file paths
            validated_config = state_manager.load_state()
            assert validated_config["file1_path"] == ""
            assert validated_config["file2_path"] == ""
            assert validated_config["output_path"] == test_config["output_path"]
            assert validated_config["key_column"] == test_config["key_column"]
            assert validated_config["key_row"] == test_config["key_row"]

            print("✅ File validation test passed!")

            # Test loading non-existent config (should return defaults)
            state_manager.config_file.unlink()  # Delete the config file
            default_config = state_manager.load_state()

            assert default_config["file1_path"] == ""
            assert default_config["file2_path"] == ""
            assert default_config["output_path"] == ""
            assert default_config["key_column"] == 1
            assert default_config["key_row"] == 1

            print("✅ Default config test passed!")


if __name__ == "__main__":
    test_state_manager()
