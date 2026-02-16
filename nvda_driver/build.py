#!/usr/bin/env python3
"""Package the SoftVoice NVDA addon as a .nvda-addon file (ZIP archive)."""

import os
import sys
import zipfile

if sys.version_info < (3, 0):
	raise Exception("Python 3 required")

ADDON_DIR = os.path.dirname(os.path.abspath(__file__))
SYNTH_DIR = os.path.join(ADDON_DIR, "synthDrivers")
OUTPUT_NAME = "softvoice.nvda-addon"
OUTPUT_PATH = os.path.join(ADDON_DIR, OUTPUT_NAME)

# Files to include in the addon
ADDON_FILES = {
	# Root addon files
	"manifest.ini": os.path.join(ADDON_DIR, "manifest.ini"),

	# Documentation
	"doc/en/readme.html": os.path.join(ADDON_DIR, "doc", "en", "readme.html"),
	"doc/en/readme.md": os.path.join(ADDON_DIR, "doc", "en", "readme.md"),
	"doc/ru/readme.html": os.path.join(ADDON_DIR, "doc", "ru", "readme.html"),
	"doc/ru/readme.md": os.path.join(ADDON_DIR, "doc", "ru", "readme.md"),
	"doc/style.css": os.path.join(ADDON_DIR, "doc", "style.css"),

	# Locale
	"locale/ru/manifest.ini": os.path.join(ADDON_DIR, "locale", "ru", "manifest.ini"),

	# SynthDriver Python files
	"synthDrivers/sv.py": os.path.join(SYNTH_DIR, "sv.py"),
	"synthDrivers/_softvoice.py": os.path.join(SYNTH_DIR, "_softvoice.py"),
	"synthDrivers/_ipc.py": os.path.join(SYNTH_DIR, "_ipc.py"),

	# 32-bit host executable
	"synthDrivers/softvoice_host32.exe": os.path.join(SYNTH_DIR, "softvoice_host32.exe"),

	# 32-bit wrapper DLL
	"synthDrivers/softvoice_wrapper.dll": os.path.join(SYNTH_DIR, "softvoice_wrapper.dll"),

	# Engine DLLs
	"synthDrivers/tibase32.dll": os.path.join(SYNTH_DIR, "tibase32.dll"),
	"synthDrivers/tieng32.dll": os.path.join(SYNTH_DIR, "tieng32.dll"),
	"synthDrivers/svctl32.dll": os.path.join(SYNTH_DIR, "svctl32.dll"),
	"synthDrivers/sveng32.dll": os.path.join(SYNTH_DIR, "sveng32.dll"),
	"synthDrivers/svctl64.dll": os.path.join(SYNTH_DIR, "svctl64.dll"),
	"synthDrivers/sveng64.dll": os.path.join(SYNTH_DIR, "sveng64.dll"),
	"synthDrivers/TISPAN32.DLL": os.path.join(SYNTH_DIR, "TISPAN32.DLL"),
}


def main():
	missing = []
	for arc_name, src_path in ADDON_FILES.items():
		if not os.path.exists(src_path):
			missing.append((arc_name, src_path))

	if missing:
		print("WARNING: Missing files:")
		for arc_name, src_path in missing:
			print(f"  {arc_name} -> {src_path}")
		print()

	# Only require the Python files and manifest to exist
	required = ["manifest.ini", "synthDrivers/sv.py", "synthDrivers/_softvoice.py",
				"synthDrivers/_ipc.py"]
	for r in required:
		if r in [m[0] for m in missing]:
			print(f"ERROR: Required file missing: {r}")
			sys.exit(1)

	print(f"Creating {OUTPUT_NAME}...")
	with zipfile.ZipFile(OUTPUT_PATH, "w", zipfile.ZIP_DEFLATED) as zf:
		for arc_name, src_path in ADDON_FILES.items():
			if os.path.exists(src_path):
				zf.write(src_path, arc_name)
				size = os.path.getsize(src_path)
				print(f"  + {arc_name} ({size:,} bytes)")
			else:
				print(f"  - {arc_name} (SKIPPED - not found)")

	print(f"\nCreated {OUTPUT_PATH}")
	print(f"Size: {os.path.getsize(OUTPUT_PATH):,} bytes")


if __name__ == "__main__":
	main()
