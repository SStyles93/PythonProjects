import sys
import fitz  # PyMuPDF
from PIL import Image, ImageFilter
import io

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QGraphicsView, QGraphicsScene,
    QGraphicsPixmapItem, QGraphicsRectItem, QSlider, QSpinBox, QMessageBox
)
from PyQt5.QtGui import QPixmap, QImage, QPen, QBrush, QColor, QIcon
from PyQt5.QtCore import Qt, QRectF

class MediaViewer(QGraphicsView):
    """A custom QGraphicsView for displaying the media and handling selection."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.scene = QGraphicsScene(self)
        self.setScene(self.scene)
        self.pixmap_item = None
        self.selection_rect = None
        self.start_pos = None

    def set_pixmap(self, pixmap):
        self.scene.clear()
        self.selection_rect = None
        self.pixmap_item = QGraphicsPixmapItem(pixmap)
        self.scene.addItem(self.pixmap_item)
        self.fitInView(self.scene.itemsBoundingRect(), Qt.AspectRatioMode.KeepAspectRatio)

    def mousePressEvent(self, event):
        if self.pixmap_item and event.button() == Qt.MouseButton.LeftButton:
            self.start_pos = self.mapToScene(event.pos())
            
            if self.selection_rect:
                self.scene.removeItem(self.selection_rect)
                self.selection_rect = None
            
            # Create a semi-transparent rectangle for selection
            pen = QPen(QColor(255, 0, 0, 200), 2, Qt.PenStyle.DashLine)
            brush = QBrush(QColor(0, 0, 255, 50))
            self.selection_rect = QGraphicsRectItem(QRectF(self.start_pos, self.start_pos))
            self.selection_rect.setPen(pen)
            self.selection_rect.setBrush(brush)
            self.scene.addItem(self.selection_rect)
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self.start_pos:
            current_pos = self.mapToScene(event.pos())
            rect = QRectF(self.start_pos, current_pos).normalized()
            self.selection_rect.setRect(rect)
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.start_pos = None
        super().mouseReleaseEvent(event)

    def get_selection(self):
        if self.selection_rect:
            return self.selection_rect.rect()
        return None

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowIcon(QIcon("icon.ico"))
        self.setWindowTitle("Media Blurring Tool")
        self.setGeometry(100, 100, 1000, 800)

        # --- Member variables ---
        self.file_path = None
        self.file_type = None # Track if it's a 'pdf' or 'image'
        self.pdf_doc = None
        self.current_page_num = 0
        self.current_pil_image = None # Store the master image (from PDF or file) as a PIL Image

        # --- Main Layout ---
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QHBoxLayout(main_widget)

        # --- Left Panel (Controls) ---
        controls_layout = QVBoxLayout()
        
        self.btn_open = QPushButton("Open Media File")
        self.btn_open.clicked.connect(self.open_file)
        
        self.page_label = QLabel("Page: N/A")
        self.btn_prev = QPushButton("<< Previous")
        self.btn_prev.clicked.connect(self.prev_page)
        self.btn_next = QPushButton("Next >>")
        self.btn_next.clicked.connect(self.next_page)
        
        page_nav_layout = QHBoxLayout()
        page_nav_layout.addWidget(self.btn_prev)
        page_nav_layout.addWidget(self.btn_next)

        self.radius_label = QLabel("Blur Radius:")
        self.radius_slider = QSlider(Qt.Orientation.Horizontal)
        self.radius_slider.setRange(1, 50)
        self.radius_slider.setValue(10)
        self.radius_spinbox = QSpinBox()
        self.radius_spinbox.setRange(1, 50)
        self.radius_spinbox.setValue(10)
        self.radius_slider.valueChanged.connect(self.radius_spinbox.setValue)
        self.radius_spinbox.valueChanged.connect(self.radius_slider.setValue)

        self.btn_blur = QPushButton("Apply Blur to Selection")
        self.btn_blur.clicked.connect(self.apply_blur)

        self.btn_save = QPushButton("Save File")
        self.btn_save.clicked.connect(self.save_file)

        controls_layout.addWidget(self.btn_open)
        controls_layout.addWidget(self.page_label)
        controls_layout.addLayout(page_nav_layout)
        controls_layout.addSpacing(20)
        controls_layout.addWidget(self.radius_label)
        controls_layout.addWidget(self.radius_slider)
        controls_layout.addWidget(self.radius_spinbox)
        controls_layout.addSpacing(20)
        controls_layout.addWidget(self.btn_blur)
        controls_layout.addStretch()
        controls_layout.addWidget(self.btn_save)

        # --- Right Panel (Viewer) ---
        self.media_viewer = MediaViewer()
        
        main_layout.addLayout(controls_layout, 1)
        main_layout.addWidget(self.media_viewer, 3)

        self.update_ui_states()

    def update_ui_states(self):
        """Enable/disable buttons based on the current state."""
        has_file = self.file_path is not None
        is_pdf = self.file_type == 'pdf'
        
        self.btn_prev.setEnabled(is_pdf and self.current_page_num > 0)
        self.btn_next.setEnabled(is_pdf and self.current_page_num < (self.pdf_doc.page_count - 1 if self.pdf_doc else 0))
        self.btn_blur.setEnabled(has_file)
        self.btn_save.setEnabled(has_file)
        
        # Hide page navigation for images
        self.page_label.setVisible(is_pdf)
        self.btn_prev.setVisible(is_pdf)
        self.btn_next.setVisible(is_pdf)

    def open_file(self):
        # Updated file dialog to accept multiple types
        file_filter = "All Supported Files (*.pdf *.png *.jpg *.jpeg *.bmp);;PDF Files (*.pdf);;Image Files (*.png *.jpg *.jpeg *.bmp)"
        path, _ = QFileDialog.getOpenFileName(self, "Open Media File", "", file_filter)
        
        if not path:
            return
            
        self.file_path = path
        try:
            if path.lower().endswith('.pdf'):
                self.file_type = 'pdf'
                self.pdf_doc = fitz.open(path)
                self.current_page_num = 0
                self.load_pdf_page()
            else:
                self.file_type = 'image'
                self.pdf_doc = None
                self.current_pil_image = Image.open(path).convert("RGB") # Ensure it's RGB
                self.display_pil_image(self.current_pil_image)
                self.page_label.setText("Image Loaded")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to open file: {e}")
            self.file_path = None
        
        self.update_ui_states()

    def display_pil_image(self, pil_img):
        """Converts a PIL image to a QPixmap and displays it."""
        buffer = io.BytesIO()
        pil_img.save(buffer, format="PNG")
        buffer.seek(0)
        
        qimage = QImage.fromData(buffer.read())
        qpixmap = QPixmap.fromImage(qimage)
        
        self.media_viewer.set_pixmap(qpixmap)

    def load_pdf_page(self):
        """Loads a specific page from a PDF, converts it to a PIL image, and displays it."""
        if not self.pdf_doc:
            return
        
        page = self.pdf_doc.load_page(self.current_page_num)
        zoom = 300 / 72  # 300 DPI for quality
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat)
        
        # Store the page as a PIL image
        self.current_pil_image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        self.display_pil_image(self.current_pil_image)
        
        self.page_label.setText(f"Page: {self.current_page_num + 1} / {self.pdf_doc.page_count}")
        self.update_ui_states()

    def prev_page(self):
        if self.file_type == 'pdf' and self.current_page_num > 0:
            self.current_page_num -= 1
            self.load_pdf_page()

    def next_page(self):
        if self.file_type == 'pdf' and self.current_page_num < self.pdf_doc.page_count - 1:
            self.current_page_num += 1
            self.load_pdf_page()

    def apply_blur(self):
        selection = self.media_viewer.get_selection()
        if not self.current_pil_image:
            QMessageBox.warning(self, "No File", "Please open a file first.")
            return
        if not selection:
            QMessageBox.warning(self, "No Selection", "Please select an area to blur by clicking and dragging.")
            return

        # The logic now operates on self.current_pil_image, which is universal
        # The selection coordinates are already scaled correctly because they come from the displayed pixmap
        box = (int(selection.left()), int(selection.top()), int(selection.right()), int(selection.bottom()))
        
        region_to_blur = self.current_pil_image.crop(box)
        radius = self.radius_slider.value()
        blurred_region = region_to_blur.filter(ImageFilter.GaussianBlur(radius=radius))
        
        # Paste the blurred region back onto the master PIL image
        self.current_pil_image.paste(blurred_region, box)
        
        # Update the view with the newly blurred image
        self.display_pil_image(self.current_pil_image)

    def save_file(self):
        if not self.file_path:
            return

        # Unified save logic
        if self.file_type == 'pdf':
            # For PDFs, we need to re-insert the modified image back into the PDF document
            page = self.pdf_doc.load_page(self.current_page_num)
            page.clean_contents()
            
            buffer = io.BytesIO()
            self.current_pil_image.save(buffer, format="PNG")
            buffer.seek(0)
            
            page.insert_image(page.rect, stream=buffer.getvalue())
            
            # Now, open the save dialog for the PDF
            path, _ = QFileDialog.getSaveFileName(self, "Save PDF", "", "PDF Files (*.pdf)")
            if path:
                try:
                    self.pdf_doc.save(path, garbage=4, deflate=True, clean=True)
                    QMessageBox.information(self, "Success", f"PDF saved successfully to {path}")
                except Exception as e:
                    QMessageBox.critical(self, "Error", f"Failed to save PDF: {e}")
        
        elif self.file_type == 'image':
            # For images, we just save the modified PIL image
            path, _ = QFileDialog.getSaveFileName(self, "Save Image", "", "PNG (*.png);;JPEG (*.jpg *.jpeg);;Bitmap (*.bmp)")
            if path:
                try:
                    self.current_pil_image.save(path)
                    QMessageBox.information(self, "Success", f"Image saved successfully to {path}")
                except Exception as e:
                    QMessageBox.critical(self, "Error", f"Failed to save image: {e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
