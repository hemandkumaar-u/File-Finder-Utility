import unittest
import os
import shutil
import tempfile
from src.file_matcher import FileMatcher
from src.file_copier import FileCopier

class TestFileMatcherAndCopier(unittest.TestCase):
    def setUp(self):
        self.source_dir = tempfile.mkdtemp()
        self.dest_dir = tempfile.mkdtemp()

        # Create dummy image files in source directory
        self.f1 = os.path.join(self.source_dir, "IMG_001.jpg")
        self.f2 = os.path.join(self.source_dir, "IMG_002.png")
        self.f3 = os.path.join(self.source_dir, "photo_xyz.png")

        with open(self.f1, "w") as f:
            f.write("fake image data 1")
        with open(self.f2, "w") as f:
            f.write("fake image data 2")
        with open(self.f3, "w") as f:
            f.write("fake image data 3")

    def tearDown(self):
        shutil.rmtree(self.source_dir, ignore_errors=True)
        shutil.rmtree(self.dest_dir, ignore_errors=True)

    def test_matching_exact_and_stem(self):
        extracted = [
            {'raw': 'IMG_001.jpg', 'id': 'img_001', 'location': 'Paragraph 1'},
            {'raw': 'IMG_002', 'id': 'img_002', 'location': 'Paragraph 2'},
            {'raw': 'MISSING_FILE', 'id': 'missing_file', 'location': 'Paragraph 3'}
        ]

        result = FileMatcher.match_docx_items(
            extracted_items=extracted,
            source_folder=self.source_dir,
            extension_agnostic=True,
            case_sensitive=False,
            image_only=False
        )

        matched = result['matched']
        missing = result['missing']

        self.assertEqual(len(matched), 2)
        self.assertEqual(len(missing), 1)

        matched_names = [m['file_name'] for m in matched]
        self.assertIn("IMG_001.jpg", matched_names)
        self.assertIn("IMG_002.png", matched_names)

    def test_file_copier(self):
        matched_items = [
            {
                'file_path': self.f1,
                'file_name': 'IMG_001.jpg',
                'id': 'img_001',
                'raw': 'IMG_001.jpg'
            }
        ]

        copy_res = FileCopier.process_files(
            selected_items=matched_items,
            dest_folder=self.dest_dir,
            overwrite=False
        )

        self.assertEqual(copy_res['copied_count'], 1)
        copied_file = os.path.join(self.dest_dir, "IMG_001.jpg")
        self.assertTrue(os.path.exists(copied_file))
        self.assertTrue(os.path.exists(copy_res['report_path']))

if __name__ == "__main__":
    unittest.main()
