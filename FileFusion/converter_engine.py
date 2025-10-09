# converter_engine.py

import os
import logging
from typing import List
from docx2pdf import convert
from PyPDF2 import PdfMerger

# Configure logging to output messages
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class ConverterEngine:
    """Handles the core logic for file conversion and merging."""

    def __init__(self, file_paths: List[str]):
        self.file_paths = file_paths
        self.output_files = []

    def convert_to_pdf(self) -> List[str]:
        """
        Converts a list of .doc and .docx files to .pdf.
        PDFs in the list are skipped but included in the final output list.
        """
        self.output_files = []
        temp_pdf_paths = []

        for i, file_path in enumerate(self.file_paths):
            filename = os.path.basename(file_path)
            output_dir = os.path.dirname(file_path)
            
            try:
                if file_path.lower().endswith(('.doc', '.docx')):
                    logging.info(f"Converting file {i+1}/{len(self.file_paths)}: {filename}")
                    
                    # Define the output path for the PDF in the same directory
                    pdf_path = os.path.join(output_dir, f"{os.path.splitext(filename)[0]}.pdf")
                    
                    # Perform the conversion
                    convert(file_path, pdf_path)
                    
                    logging.info(f"Successfully converted {filename} to PDF.")
                    self.output_files.append(pdf_path)
                    temp_pdf_paths.append(pdf_path) # Keep track of created PDFs for cleanup

                elif file_path.lower().endswith('.pdf'):
                    logging.info(f"File {i+1}/{len(self.file_paths)}: {filename} is already a PDF. Skipping conversion.")
                    self.output_files.append(file_path)
                
                else:
                    logging.warning(f"Unsupported file type: {filename}. Skipping.")

            except Exception as e:
                logging.error(f"Failed to process {filename}. Error: {e}")
                # Clean up any partially created PDFs if conversion fails
                self._cleanup(temp_pdf_paths)
                raise # Re-raise the exception to be caught by the controller

        return self.output_files

    def merge_pdfs(self, output_filename: str) -> bool:
        """
        Merges the list of generated/provided PDFs into a single file.
        
        Args:
            output_filename (str): The path for the final merged PDF.
        """
        if not self.output_files:
            logging.error("No PDF files available to merge.")
            return False

        merger = PdfMerger()
        logging.info("Starting PDF merge process...")

        try:
            for pdf_path in self.output_files:
                logging.info(f"Appending {os.path.basename(pdf_path)} to merge list.")
                merger.append(pdf_path)
            
            logging.info(f"Writing merged PDF to {output_filename}...")
            merger.write(output_filename)
            merger.close()
            
            logging.info("Merge complete.")
            return True
        except Exception as e:
            logging.error(f"Failed to merge PDFs. Error: {e}")
            return False
        finally:
            # Clean up intermediate PDFs created during conversion if merging was the goal
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

