# converter_engine.py (Corrected for Output Folder)

import os
import logging
from typing import List, Callable 
from docx2pdf import convert
from PyPDF2 import PdfMerger
from PIL import Image

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class ConverterEngine:
    """Handles the core logic for file conversion and merging."""

    image_extensions = ('.png', '.jpg', '.jpeg', '.bmp', '.tiff')

    def __init__(self, file_paths: List[str]):
        self.file_paths = file_paths
        self.output_files = []

    # --- The method accepts an optional 'output_folder' argument ---
    def convert_to_pdf(self, output_folder=None, progress_callback: Callable[[int], None] = None) -> List[str]:
        """
        Converts documents and images to PDF, reporting progress after each file.

        Args:
            output_folder (str, optional): The folder to save converted files in.
            progress_callback (Callable, optional): A function to call with the new progress percentage.
        """
        self.output_files = []
        temp_pdf_paths = []
        total_files = len(self.file_paths)

        for i, file_path in enumerate(self.file_paths):
            filename = os.path.basename(file_path)
            file_ext = os.path.splitext(filename)[1].lower()
            
            # Determine the output directory based on whether 'output_folder' was provided.
            if output_folder:
                output_dir = output_folder
            else:
                # Fallback for the merge workflow (where temp files are created)
                output_dir = os.path.dirname(file_path)
            
            try:
                if file_ext in ('.doc', '.docx') or file_ext in self.image_extensions:
                    pdf_path = os.path.join(output_dir, f"{os.path.splitext(filename)[0]}.pdf")
                    if file_ext in ('.doc', '.docx'):
                        convert(file_path, pdf_path)
                    elif file_ext in self.image_extensions:
                        image = Image.open(file_path)
                        if image.mode == 'RGBA': image = image.convert('RGB')
                        image.save(pdf_path, "PDF", resolution=100.0)
                    
                    logging.info(f"Successfully converted {filename} to PDF.")
                    self.output_files.append(pdf_path)
                    if not output_folder: temp_pdf_paths.append(pdf_path)

                elif file_ext == '.pdf':
                    logging.info(f"File {i+1}/{total_files}: {filename} is already a PDF. Skipping.")
                    self.output_files.append(file_path)
                
                else:
                    logging.warning(f"Unsupported file type: {filename}. Skipping.")

                # After successfully processing a file, call the callback.
                if progress_callback:
                    # Calculate the percentage based on the user's formula.
                    # We use (i + 1) because the loop is 0-indexed.
                    percentage = int(((i + 1) / total_files) * 100)
                    progress_callback(percentage)

            except Exception as e:
                logging.error(f"Failed to process {filename}. Error: {e}")
                self._cleanup(temp_pdf_paths)
                raise

        return self.output_files

    def merge_pdfs(self, output_file_object):
        """
        Merges the list of generated/provided PDFs into a single file object.
        """
        if not self.output_files:
            logging.error("No PDF files available to merge.")
            return False

        merger = PdfMerger()
        logging.info("Starting PDF merge process...")

        try:
            for pdf_path in self.output_files:
                merger.append(pdf_path)
            
            merger.write(output_file_object)
            merger.close()
            
            logging.info("Merge complete.")
            return True
        except Exception as e:
            logging.error(f"Failed to merge PDFs. Error: {e}")
            return False
        finally:
            # This cleanup is now primarily for the intermediate files created during a merge.
            temp_pdfs = [f for f in self.output_files if f not in self.file_paths]
            self._cleanup(temp_pdfs)

    def _cleanup(self, file_paths: List[str]):
        """Removes temporary files."""
        for path in file_paths:
            try:
                os.remove(path)
                logging.info(f"Cleaned up temporary file: {path}")
            except OSError as e:
                logging.error(f"Error cleaning up file {path}: {e}")
