import os
import glob
from typing import List, Dict, Any, Optional

IMAGE_EXTENSIONS = ('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp', '.tiff', '.tif', '.svg', '.ico', '.heic', '.raw')

class FileMatcher:
    """Matches IDs extracted from Word documents against files present in a folder."""

    @staticmethod
    def scan_folder(folder_path: str, recursive: bool = True, image_only: bool = True) -> List[str]:
        """Scans folder for files and returns list of absolute file paths."""
        if not os.path.exists(folder_path):
            raise FileNotFoundError(f"Source folder not found: {folder_path}")

        file_paths = []
        if recursive:
            for root, _, files in os.walk(folder_path):
                for f in files:
                    full_path = os.path.join(root, f)
                    if not image_only or f.lower().endswith(IMAGE_EXTENSIONS):
                        file_paths.append(full_path)
        else:
            for f in os.listdir(folder_path):
                full_path = os.path.join(folder_path, f)
                if os.path.isfile(full_path):
                    if not image_only or f.lower().endswith(IMAGE_EXTENSIONS):
                        file_paths.append(full_path)

        return file_paths

    @classmethod
    def match_docx_items(
        cls,
        extracted_items: List[Dict[str, Any]],
        source_folder: str,
        recursive: bool = True,
        case_sensitive: bool = False,
        extension_agnostic: bool = True,
        image_only: bool = True
    ) -> Dict[str, Any]:
        """
        Matches docx extracted items against source folder files.
        Returns a dict:
        {
          'matched': [ {...file_path, file_name, file_size, id, raw, location, selected: True...} ],
          'missing': [ {...raw, id, location...} ]
        }
        """
        all_files = cls.scan_folder(source_folder, recursive=recursive, image_only=image_only)

        # Build index maps for fast lookups
        # 1. Full exact filename map: "image.png" -> list of paths
        # 2. Stem map: "image" -> list of paths
        exact_map: Dict[str, List[str]] = {}
        stem_map: Dict[str, List[str]] = {}

        for filepath in all_files:
            fname = os.path.basename(filepath)
            stem, _ = os.path.splitext(fname)

            key_exact = fname if case_sensitive else fname.lower()
            key_stem = stem if case_sensitive else stem.lower()

            exact_map.setdefault(key_exact, []).append(filepath)
            stem_map.setdefault(key_stem, []).append(filepath)

        matched_results = []
        matched_item_keys = set()
        missing_items = []

        for item in extracted_items:
            raw_text = item['raw']
            item_id = item['id']
            location = item['location']

            key_raw = raw_text if case_sensitive else raw_text.lower()
            key_stem = item_id if case_sensitive else item_id.lower()

            found_paths = []
            match_type = ""

            # Check exact filename match first
            if key_raw in exact_map:
                found_paths = exact_map[key_raw]
                match_type = "Exact Filename Match"
            elif key_stem in exact_map:
                found_paths = exact_map[key_stem]
                match_type = "Exact Match"
            # Check stem match (extension agnostic)
            elif extension_agnostic and key_stem in stem_map:
                found_paths = stem_map[key_stem]
                match_type = "Extension-Agnostic Match"
            else:
                # Try finding if item_id matches filename substring or sanitized name
                sanitized_stem = key_stem.replace('_', ' ').replace('-', ' ').strip()
                for stem_key, paths in stem_map.items():
                    clean_key = stem_key.replace('_', ' ').replace('-', ' ').strip()
                    if clean_key == sanitized_stem:
                        found_paths = paths
                        match_type = "Flexible Format Match"
                        break

            if found_paths:
                for fpath in found_paths:
                    file_key = (fpath.lower(), location.lower())
                    if file_key not in matched_item_keys:
                        matched_item_keys.add(file_key)
                        file_size = os.path.getsize(fpath) if os.path.exists(fpath) else 0
                        matched_results.append({
                            'id': item_id,
                            'raw': raw_text,
                            'location': location,
                            'file_path': fpath,
                            'file_name': os.path.basename(fpath),
                            'file_size': file_size,
                            'extension': os.path.splitext(fpath)[1].lower(),
                            'match_type': match_type,
                            'selected': True
                        })
            else:
                missing_items.append({
                    'raw': raw_text,
                    'id': item_id,
                    'location': location
                })

        return {
            'matched': matched_results,
            'missing': missing_items,
            'total_extracted': len(extracted_items),
            'total_folder_files': len(all_files)
        }
