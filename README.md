# File-Finder-Utility
File Finder Utility allows a user to provide a Microsoft Word document (.docx) containing a list of image names or file IDs. The application extracts all the requested IDs, searches a specified local directory for files matching those IDs, and automatically copies them into a output folder.


Key Features:

Smart Document Parsing: Automatically extracts file names and IDs from Word documents, intelligently navigating paragraphs, bulleted lists, and tables. It handles prefixes (e.g., "Image ID: 123") and comma-separated lists smoothly.
Automated File Matching: Scans a source directory to find the actual files that correspond to the IDs found in the Word document, agnostic of specific image extensions.
Bulk File Extraction: Copies all successfully matched files into a target destination folder, saving hours of manual searching and dragging-and-dropping.
Graphical User Interface (GUI): Features a user-friendly desktop interface allowing non-technical users to easily select their input document, source folder, and output destination.
Standalone Application: Configured with PyInstaller (.spec file included) to be packaged as a standalone executable, meaning users can run it without needing to install Python or any dependencies.
