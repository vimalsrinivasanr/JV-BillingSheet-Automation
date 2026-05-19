import os
import pytest

def test_input_folder_writable():
    test_file = os.path.join('input', 'test_write.txt')
    try:
        with open(test_file, 'w') as f:
            f.write('test')
        assert os.path.exists(test_file)
    finally:
        if os.path.exists(test_file):
            os.remove(test_file)

def test_output_folder_writable():
    test_file = os.path.join('output', 'test_write.txt')
    try:
        with open(test_file, 'w') as f:
            f.write('test')
        assert os.path.exists(test_file)
    finally:
        if os.path.exists(test_file):
            os.remove(test_file)
