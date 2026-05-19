import pytest
import os

def test_input_folder_exists():
    assert os.path.exists('input')

def test_normalized_folder_exists():
    assert os.path.exists('normalized')

def test_output_folder_exists():
    assert os.path.exists('output')

def test_scripts_folder_exists():
    assert os.path.exists('scripts')
