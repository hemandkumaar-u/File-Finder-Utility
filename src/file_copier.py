import os
import shutil
import csv
from datetime import datetime
from typing import List, Dict, Any, Callable, Optional

class FileCopier:
    """Handles copying or moving matched image files to the destination directory and generating reports."""

    @staticmethod
    def get_unique_filepath(dest_folder: str, filename: str) -> str:
        """Generates a non-conflicting filepath if file already exists (e.g., photo (1).jpg)."""
        base_path = os.path.join(dest_folder, filename)
        if not os.path.exists(base_path):
            return base_path

        stem, ext = os.path.splitext(filename)
        counter = 1
        while True:
            new_name = f"{stem} ({counter}){ext}"
            new_path = os.path.join(dest_folder, new_name)
            if not os.path.exists(new_path):
                return new_path
            counter += 1

    @classmethod
    def process_files(
        cls,
        selected_items: List[Dict[str, Any]],
        dest_folder: str,
        missing_items: Optional[List[Dict[str, Any]]] = None,
        overwrite: bool = False,
        move_files: bool = False,
        progress_callback: Optional[Callable[[int, int, str], None]] = None
    ) -> Dict[str, Any]:
        """
        Copies or moves selected files to dest_folder.
        Optionally generates a summary CSV report in dest_folder.
        """
        if not os.path.exists(dest_folder):
            os.makedirs(dest_folder, exist_ok=True)

        copied_count = 0
        failed_count = 0
        skipped_count = 0
        copied_details = []

        total = len(selected_items)

        for idx, item in enumerate(selected_items):
            src_path = item['file_path']
            filename = item['file_name']

            if not os.path.exists(src_path):
                failed_count += 1
                if progress_callback:
                    progress_callback(idx + 1, total, f"Source missing: {filename}")
                continue

            if overwrite:
                target_path = os.path.join(dest_folder, filename)
            else:
                target_path = cls.get_unique_filepath(dest_folder, filename)

            try:
                if move_files:
                    shutil.move(src_path, target_path)
                    action_done = "Moved"
                else:
                    shutil.copy2(src_path, target_path)
                    action_done = "Copied"

                copied_count += 1
                copied_details.append({
                    'id': item.get('id', ''),
                    'word_text': item.get('raw', ''),
                    'source_path': src_path,
                    'destination_path': target_path,
                    'status': action_done,
                    'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                })
                if progress_callback:
                    progress_callback(idx + 1, total, f"{action_done}: {os.path.basename(target_path)}")
            except Exception as e:
                failed_count += 1
                if progress_callback:
                    progress_callback(idx + 1, total, f"Error processing {filename}: {str(e)}")

        # Save CSV summary report in destination folder
        report_path = os.path.join(dest_folder, "matched_files_summary.csv")
        try:
            with open(report_path, mode='w', newline='', encoding='utf-8') as csvfile:
                fieldnames = ['ID', 'Word Text', 'Match Status', 'Source File Path', 'Destination File Path', 'Timestamp']
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                writer.writeheader()

                for d in copied_details:
                    writer.writerow({
                        'ID': d['id'],
                        'Word Text': d['word_text'],
                        'Match Status': d['status'],
                        'Source File Path': d['source_path'],
                        'Destination File Path': d['destination_path'],
                        'Timestamp': d['timestamp']
                    })

                if missing_items:
                    for m in missing_items:
                        writer.writerow({
                            'ID': m.get('id', ''),
                            'Word Text': m.get('raw', ''),
                            'Match Status': 'NOT FOUND IN SOURCE FOLDER',
                            'Source File Path': 'N/A',
                            'Destination File Path': 'N/A',
                            'Timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        })
        except Exception:
            report_path = ""

        return {
            'copied_count': copied_count,
            'failed_count': failed_count,
            'skipped_count': skipped_count,
            'report_path': report_path,
            'copied_details': copied_details
        }
