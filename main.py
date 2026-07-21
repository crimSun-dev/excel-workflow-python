"""Top-level entry point.

    python main.py                          -> launches the GUI file picker
    python main.py process -r ... -f ...     -> runs the CLI pipeline

Also serves as the PyInstaller build target.
"""

from src.cli import main

if __name__ == "__main__":
    main()
