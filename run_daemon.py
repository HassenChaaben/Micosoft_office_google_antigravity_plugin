"""
Stand-alone entrypoint for PyInstaller packaging.
"""
import os
import sys

# Ensure project root is in sys.path
root_dir = os.path.dirname(os.path.abspath(__file__))
if root_dir not in sys.path:
    sys.path.insert(0, root_dir)

from daemon.main import main

if __name__ == "__main__":
    main()
