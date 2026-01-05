
import pypandoc
import os
import sys

print("Checking Pandoc presence...")
try:
    pypandoc.get_pandoc_path()
    print("Pandoc is already available.")
except OSError:
    print("Pandoc not found. Attempting to download...")
    try:
        pypandoc.download_pandoc()
        print("Pandoc download successful.")
    except Exception as e:
        print(f"Failed to download Pandoc: {e}")
        sys.exit(1)
