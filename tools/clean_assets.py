#!/usr/bin/env python3

import os
import argparse
from pathlib import Path

def find_images_in_qmd(qmd_content):
    """
    Very basic extractor to find anything like `assets/basename/filename.ext`.
    Quarto images could be `![](assets/...)` or `<img src="assets/..."/>`
    """
    import re
    # Match both standard markdown and HTML src attributes
    pattern = re.compile(r'(?:\]\()(?P<md>assets/[^)]+)(?:\))|(?:src=")(?P<html>assets/[^"]+)(?:")')
    images = set()
    for match in pattern.finditer(qmd_content):
        if match.group('md'):
            images.add(match.group('md'))
        elif match.group('html'):
            images.add(match.group('html'))
    return images

def clean_assets(present_dir="present", dry_run=False):
    present_path = Path(present_dir)
    assets_dir = present_path / "assets"
    
    if not present_path.exists() or not assets_dir.exists():
        print(f"Directory structure {present_dir}/assets not found. Exiting.")
        return

    # Find all QMD files
    qmd_files = list(present_path.glob("*.qmd"))
    if not qmd_files:
        print(f"No .qmd files found in {present_dir}.")
        return

    # Collect all referenced images across all QMDs
    referenced_assets = set()
    for qmd in qmd_files:
        try:
            content = qmd.read_text(encoding='utf-8')
            refs = find_images_in_qmd(content)
            # Normalize paths to handle relative logic safely
            for ref in refs:
                # Assuming the ref is "assets/basename/image.png" relative to present/
                normalized_ref = str(Path(ref))
                referenced_assets.add(normalized_ref)
        except Exception as e:
            print(f"Error reading {qmd.name}: {e}")

    # Now iterate through the actual assets directory and see what exists vs what is referenced
    total_deleted = 0
    total_saved = 0

    print("========================================")
    print("Orphaned Asset Scan")
    print("========================================")

    # Walk through all files in the assets directory
    for root, _, files in os.walk(assets_dir):
        for file in files:
            file_path = Path(root) / file
            
            # Reconstruct the relative path as it would appear in the markdown: "assets/.../..."
            try:
                # relative_to present_path gives us "assets/subfolder/img.png"
                rel_path = str(file_path.relative_to(present_path))
            except ValueError:
                continue

            if rel_path not in referenced_assets:
                file_size = file_path.stat().st_size
                if dry_run:
                    print(f"[DRY-RUN] Will delete: {rel_path} ({(file_size/1024):.1f} KB)")
                else:
                    print(f"Deleting: {rel_path} ({(file_size/1024):.1f} KB)")
                    file_path.unlink()
                
                total_deleted += 1
                total_saved += file_size
            else:
                # Keep it
                pass

    if total_deleted == 0:
        print("No orphaned assets found. Everything is clean!")
    else:
        status = "Would have freed" if dry_run else "Freed"
        print(f"\nFinished: {total_deleted} files processed. {status} {(total_saved/1024/1024):.2f} MB.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Clean up unused assets from Quarto presentations.")
    parser.add_argument("--dry-run", action="store_true", help="Print files that would be deleted without actually deleting them.")
    parser.add_argument("--dir", default="present", help="Directory containing the .qmd files (default: present)")
    args = parser.parse_args()

    clean_assets(args.dir, args.dry_run)
