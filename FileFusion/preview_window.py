import fitz  # The PyMuPDF library
import os 
import gc
import logging

from PyQt5.QtWidgets import QDialog, QVBoxLayout, QPushButton, QHBoxLayout, QLabel, QScrollArea
from PyQt5.QtGui import QPixmap, QImage
from PyQt5.QtCore import pyqtSignal, Qt

class PreviewWindow(QDialog):
    """
    A dialog window to preview a PDF using PyMuPDF for rendering.
    """
    closing = pyqtSignal()

    def __init__(self, temp_pdf_path, parent=None):
        super().__init__(parent)
        self.temp_pdf_path = temp_pdf_path
        
        if not os.path.exists(self.temp_pdf_path):
            raise FileNotFoundError(f"Temporary PDF file not found at {self.temp_pdf_path}")

        self.doc = fitz.open(self.temp_pdf_path)
        self.current_page = 0
        self.zoom_level = 1.0  # 1.0 = 100% zoom

        self.setWindowTitle("PDF Preview")
        self.setGeometry(150, 150, 800, 900)

        # --- Main Layout ---
        layout = QVBoxLayout(self)

        # --- Page Display ---
        self.page_label = QLabel("Loading preview...")
        self.page_label.setAlignment(Qt.AlignCenter)
        
        # --- Scroll Area for Zooming/Panning ---
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidget(self.page_label)
        self.scroll_area.setWidgetResizable(True)
        layout.addWidget(self.scroll_area)

        # --- Navigation and Info Layout ---
        nav_layout = QHBoxLayout()
        self.prev_button = QPushButton("< Previous")
        self.page_info_label = QLabel()
        self.next_button = QPushButton("Next >")
        
        nav_layout.addWidget(self.prev_button)
        nav_layout.addStretch()
        nav_layout.addWidget(self.page_info_label)
        nav_layout.addStretch()
        nav_layout.addWidget(self.next_button)
        layout.addLayout(nav_layout)

        # --- Zoom Layout ---
        zoom_layout = QHBoxLayout()
        self.zoom_in_button = QPushButton("Zoom In (+)")
        self.zoom_out_button = QPushButton("Zoom Out (-)")
        zoom_layout.addStretch()
        zoom_layout.addWidget(self.zoom_out_button)
        zoom_layout.addWidget(self.zoom_in_button)
        zoom_layout.addStretch()
        layout.addLayout(zoom_layout)

        # --- Save/Cancel Button Layout ---
        dialog_buttons_layout = QHBoxLayout()
        self.save_button = QPushButton("Save As...")
        self.cancel_button = QPushButton("Cancel")
        dialog_buttons_layout.addStretch()
        dialog_buttons_layout.addWidget(self.cancel_button)
        dialog_buttons_layout.addWidget(self.save_button)
        dialog_buttons_layout.addStretch()
        layout.addLayout(dialog_buttons_layout)

        # --- Connect Signals ---
        self.prev_button.clicked.connect(self.show_previous_page)
        self.next_button.clicked.connect(self.show_next_page)
        self.zoom_in_button.clicked.connect(self.zoom_in)
        self.zoom_out_button.clicked.connect(self.zoom_out)
        self.save_button.clicked.connect(self.accept)
        self.cancel_button.clicked.connect(self.reject)

        # --- Initial Render ---
        self.render_page()

    def render_page(self):
        """Renders the current page to the QLabel."""
        # Disable buttons at limits
        self.prev_button.setEnabled(self.current_page > 0)
        self.next_button.setEnabled(self.current_page < self.doc.page_count - 1)
        self.page_info_label.setText(f"Page {self.current_page + 1} of {self.doc.page_count}")

        # Render page using PyMuPDF
        page = self.doc.load_page(self.current_page)
        
        # Create a matrix for zooming
        mat = fitz.Matrix(self.zoom_level * 2, self.zoom_level * 2) # Use a factor of 2 for higher DPI
        pix = page.get_pixmap(matrix=mat)
        
        # Convert to QImage
        image_format = QImage.Format_RGB888 if pix.alpha == 0 else QImage.Format_RGBA8888
        qimage = QImage(pix.samples, pix.width, pix.height, pix.stride, image_format)
        
        pixmap = QPixmap.fromImage(qimage)
        self.page_label.setPixmap(pixmap)

    def show_previous_page(self):
        if self.current_page > 0:
            self.current_page -= 1
            self.render_page()

    def show_next_page(self):
        if self.current_page < self.doc.page_count - 1:
            self.current_page += 1
            self.render_page()

    def zoom_in(self):
        self.zoom_level *= 1.2
        self.render_page()

    def zoom_out(self):
        self.zoom_level /= 1.2
        self.render_page()

    def closeEvent(self, event):
        """
        Aggressively clean up PyMuPDF resources to ensure the file lock is released.
        """
        logging.info("Preview window closing. Releasing file lock...")
        
        # 1. Explicitly close the document handle from PyMuPDF.
        if self.doc:
            self.doc.close()
            logging.info("PyMuPDF document closed.")

        # 2. Delete the Python reference to the object.
        del self.doc
        self.doc = None
        
        # 3. Suggest to Python's garbage collector that now is a good time to clean up.
        # This helps destroy the underlying C++ objects holding the file handle.
        gc.collect()
        logging.info("Garbage collection triggered.")

        # 4. Finally, emit the signal for the main app to perform the file deletion.
        self.closing.emit()
        super().closeEvent(event)

