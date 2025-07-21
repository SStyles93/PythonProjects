import sys
import fitz  # PyMuPDF
from PIL import Image, ImageFilter
import io

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QGraphicsView, QGraphicsScene,
    QGraphicsPixmapItem, QGraphicsRectItem, QSlider, QSpinBox, QMessageBox
)
from PyQt5.QtGui import QPixmap, QImage, QPen, QBrush, QColor
from PyQt5.QtCore import Qt, QRectF

class PDFViewer(QGraphicsView):
    """A custom QGraphicsView for displaying the PDF and handling selection."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.scene = QGraphicsScene(self)
        self.setScene(self.scene)
        self.pixmap_item = None
        self.selection_rect = None
        self.start_pos = None

    def set_page_pixmap(self, pixmap):
        self.scene.clear()
        self.pixmap_item = QGraphicsPixmapItem(pixmap)
        self.scene.addItem(self.pixmap_item)
        self.fitInView(self.scene.itemsBoundingRect(), Qt.AspectRatioMode.KeepAspectRatio)

    def mousePressEvent(self, event):
        if self.pixmap_item and event.button() == Qt.MouseButton.LeftButton:
            self.start_pos = self.mapToScene(event.pos())
            if self.selection_rect:
                self.scene.removeItem(self.selection_rect)
            
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
        self.setWindowTitle("PDF Blurring Tool")
        self.setGeometry(100, 100, 1000, 800)

        # --- Member variables ---
        self.pdf_doc = None
        self.current_page_num = 0
        self.current_pixmap = None
        self.zoom_factor = 300 / 72  # 300 DPI for high quality

        # --- Main Layout ---
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QHBoxLayout(main_widget)

        # --- Left Panel (Controls) ---
        controls_layout = QVBoxLayout()
        
        self.btn_open = QPushButton("Open PDF")
        self.btn_open.clicked.connect(self.open_pdf)
        
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

        self.btn_save = QPushButton("Save PDF")
        self.btn_save.clicked.connect(self.save_pdf)

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

        # --- Right Panel (PDF Viewer) ---
        self.pdf_viewer = PDFViewer()
        
        main_layout.addLayout(controls_layout, 1) # 1/4 of the space
        main_layout.addWidget(self.pdf_viewer, 3) # 3/4 of the space

        self.update_button_states()

    def update_button_states(self):
        has_doc = self.pdf_doc is not None
        self.btn_prev.setEnabled(has_doc and self.current_page_num > 0)
        self.btn_next.setEnabled(has_doc and self.current_page_num < self.pdf_doc.page_count - 1)
        self.btn_blur.setEnabled(has_doc)
        self.btn_save.setEnabled(has_doc)

    def open_pdf(self):
        path, _ = QFileDialog.getOpenFileName(self, "Open PDF", "", "PDF Files (*.pdf)")
        if path:
            try:
                self.pdf_doc = fitz.open(path)
                self.current_page_num = 0
                self.load_page()
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to open PDF: {e}")

    def load_page(self):
        if not self.pdf_doc:
            return
        
        page = self.pdf_doc.load_page(self.current_page_num)
        mat = fitz.Matrix(self.zoom_factor, self.zoom_factor)
        pix = page.get_pixmap(matrix=mat)
        
        self.current_pixmap = pix # Store the original pixmap
        
        image = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format.Format_RGB888)
        qpixmap = QPixmap.fromImage(image)
        
        self.pdf_viewer.set_page_pixmap(qpixmap)
        self.page_label.setText(f"Page: {self.current_page_num + 1} / {self.pdf_doc.page_count}")
        self.update_button_states()

    def prev_page(self):
        if self.current_page_num > 0:
            self.current_page_num -= 1
            self.load_page()

    def next_page(self):
        if self.current_page_num < self.pdf_doc.page_count - 1:
            self.current_page_num += 1
            self.load_page()

    def apply_blur(self):
        selection = self.pdf_viewer.get_selection()
        if not selection:
            QMessageBox.warning(self, "No Selection", "Please select an area to blur by clicking and dragging.")
            return

        # Convert the original pixmap to a PIL Image
        img = Image.frombytes("RGB", [self.current_pixmap.width, self.current_pixmap.height], self.current_pixmap.samples)

        # The selection coordinates are already in the zoomed image space
        box = (int(selection.left()), int(selection.top()), int(selection.right()), int(selection.bottom()))
        
        region_to_blur = img.crop(box)
        radius = self.radius_slider.value()
        blurred_region = region_to_blur.filter(ImageFilter.GaussianBlur(radius=radius))
        img.paste(blurred_region, box)

        # Convert the blurred PIL Image back to a QPixmap to display it
        buffer = io.BytesIO()
        img.save(buffer, format="PNG")
        buffer.seek(0)
        
        qimage = QImage.fromData(buffer.read())
        qpixmap = QPixmap.fromImage(qimage)
        
        # Update the view with the blurred image
        self.pdf_viewer.set_page_pixmap(qpixmap)
        
        # IMPORTANT: Replace the page content in the PyMuPDF document object
        page = self.pdf_doc.load_page(self.current_page_num)
        page.clean_contents() # Clean old content
        
        # We need to convert the PIL image back to a pixmap for insertion
        # This is a bit tricky, let's use a buffer
        buffer.seek(0)
        page.insert_image(page.rect, stream=buffer.getvalue())

    def save_pdf(self):
        path, _ = QFileDialog.getSaveFileName(self, "Save PDF", "", "PDF Files (*.pdf)")
        if path:
            try:
                self.pdf_doc.save(path, garbage=4, deflate=True, clean=True)
                QMessageBox.information(self, "Success", f"PDF saved successfully to {path}")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to save PDF: {e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
