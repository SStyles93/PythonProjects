import sys
import random
import math
import time
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QFormLayout, QLabel, QSpinBox, QPushButton, QCheckBox, QSlider,
    QComboBox
)
from PyQt5.QtGui import QPainter, QColor
from PyQt5.QtCore import Qt

# --- Character Representations ---
WALL_CHAR = "█"
PATH_CHAR = " "
START_CHAR = "S"
END_CHAR = "E"

# --- New "Multiple Solutions Weaving" Logic ---
class PathGenerator:
    def __init__(self, start_pos, end_pos, spacing, branching, fill_percent, seed):
        self.user_start_pos = start_pos
        self.user_end_pos = end_pos
        self.spacing = spacing
        self.step = spacing + 1
        self.allow_branching = branching
        self.fill_percent = fill_percent
        self.seed = seed

        # Calculate grid size
        max_x = max(start_pos[0], end_pos[0])
        max_z = max(start_pos[1], end_pos[1])
        required_width = max_x + self.step + 2
        required_height = max_z + self.step + 2
        
        self.width = self._adjust_dim(required_width)
        self.height = self._adjust_dim(required_height)

        # Sanitize points to fit the final grid
        self.start_pos = self._sanitize_coord(self.user_start_pos)
        self.end_pos = self._sanitize_coord(self.user_end_pos)

        # Initialize grid
        self.grid = [[WALL_CHAR for _ in range(self.width)] for _ in range(self.height)]

    def _adjust_dim(self, dim):
        k = math.ceil((dim - 1) / self.step)
        return int(k * self.step + 1)

    def _sanitize_coord(self, pos):
        x, z = pos
        kx = round((x - 1) / self.step)
        kz = round((z - 1) / self.step)
        valid_x = int(kx * self.step + 1)
        valid_z = int(kz * self.step + 1)
        valid_x = max(1, min(valid_x, self.width - 2))
        valid_z = max(1, min(valid_z, self.height - 2))
        return (valid_x, valid_z)

    def generate(self):
        """Main generation function using the new weaving algorithm."""
        if self.seed is not None and self.seed != -1:
            random.seed(self.seed)

        # Stage 1: Determine how many paths to generate based on fill percentage
        # Map fill_percent (1-100) to a number of paths (e.g., 1-15)
        min_paths = 1
        max_paths = 15
        num_paths_to_generate = round(min_paths + (max_paths - min_paths) * (self.fill_percent / 100.0))

        # Stage 2: Generate and weave multiple solution paths
        all_carved_cells = set()
        for _ in range(num_paths_to_generate):
            path = self._generate_random_solution_path()
            all_carved_cells.update(path)
            for x, z, px, pz in path:
                # Carve the path from the previous cell (px, pz) to the current one (x, z)
                dx, dz = x - px, z - pz
                for i in range(self.step + 1):
                    ix = px + (dx // self.step) * i
                    iz = pz + (dz // self.step) * i
                    self.grid[iz][ix] = PATH_CHAR
        
        # Stage 3: Optional dead-end branching
        if self.allow_branching:
            # We can run a limited expansion to add some dead ends
            self._add_dead_end_branches(list(all_carved_cells))

        # Final step: Place markers
        self.grid[self.start_pos[1]][self.start_pos[0]] = START_CHAR
        self.grid[self.end_pos[1]][self.end_pos[0]] = END_CHAR
        
        return self.grid

    def _generate_random_solution_path(self):
        """
        Performs a biased random walk from start to end.
        Returns a list of tuples (x, z, prev_x, prev_z) representing path segments.
        """
        path_segments = []
        current = self.start_pos
        prev = self.start_pos

        # Limit steps to prevent infinite loops if it gets stuck
        for _ in range(self.width * self.height): 
            if current == self.end_pos:
                break

            # Biased walk: prefer moves that get closer to the end point
            dx_total = self.end_pos[0] - current[0]
            dz_total = self.end_pos[1] - current[1]

            moves = []
            if dx_total > 0: moves.extend([(self.step, 0)] * 4)
            elif dx_total < 0: moves.extend([(-self.step, 0)] * 4)
            if dz_total > 0: moves.extend([(0, self.step)] * 4)
            elif dz_total < 0: moves.extend([(0, -self.step)] * 4)
            
            # Add some random moves to make the path less straight
            moves.extend([(self.step, 0), (-self.step, 0), (0, self.step), (0, -self.step)])
            random.shuffle(moves)

            moved = False
            for move_dx, move_dz in moves:
                next_pos = (current[0] + move_dx, current[1] + move_dz)
                # Check if the target cell is within the grid
                if 0 < next_pos[0] < self.width - 1 and 0 < next_pos[1] < self.height - 1:
                    prev = current
                    current = next_pos
                    path_segments.append((current[0], current[1], prev[0], prev[1]))
                    moved = True
                    break
            
            if not moved: break # Exit if stuck
        
        return path_segments

    def _add_dead_end_branches(self, initial_cells):
        """Adds traditional dead-end branches growing from the main path network."""
        # Calculate how many dead-end cells to add (e.g., 20% of the main network size)
        cells_to_add = int(len(initial_cells) * 0.2) 
        
        frontier = list(initial_cells)
        random.shuffle(frontier)

        while cells_to_add > 0 and frontier:
            current = frontier.pop()
            x, z = current[0], current[1] # Handle cells that might not be full tuples
            
            neighbors = []
            for dx, dz in [(0, self.step), (0, -self.step), (self.step, 0), (-self.step, 0)]:
                nx, nz = x + dx, z + dz
                if 0 < nx < self.width - 1 and 0 < nz < self.height - 1 and self.grid[nz][nx] == WALL_CHAR:
                    neighbors.append((nx, nz))
            
            if neighbors:
                next_cell = random.choice(neighbors)
                nx, nz = next_cell
                
                dx_carve, dz_carve = nx - x, nz - z
                for i in range(self.step + 1):
                    ix = x + (dx_carve // self.step) * i
                    iz = z + (dz_carve // self.step) * i
                    if self.grid[iz][ix] == WALL_CHAR:
                        self.grid[iz][ix] = PATH_CHAR
                        cells_to_add -= 1
                
                frontier.append(next_cell)

# --- PyQt5 GUI (Unchanged from previous version) ---
class GridWidget(QWidget):
    def __init__(self):
        super().__init__(); self.grid_data = []; self.setMinimumSize(400, 400)
        self.colors = {WALL_CHAR: QColor(50, 50, 50), PATH_CHAR: QColor(220, 220, 220), START_CHAR: QColor(0, 180, 0), END_CHAR: QColor(180, 0, 0)}
    def set_grid_data(self, grid_data): self.grid_data = grid_data; self.update()
    def paintEvent(self, event):
        if not self.grid_data: return
        painter = QPainter(self); grid_h, grid_w = len(self.grid_data), len(self.grid_data[0])
        cell_w, cell_h = self.width() / grid_w, self.height() / grid_h
        for z, row in enumerate(self.grid_data):
            for x, cell_type in enumerate(row):
                painter.fillRect(int(x * cell_w), int(z * cell_h), math.ceil(cell_w), math.ceil(cell_h), self.colors.get(cell_type))

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__(); self.setWindowTitle("Multiple Solutions Weaver"); self.setGeometry(100, 100, 850, 600)
        central_widget = QWidget(); self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)
        controls_panel = QWidget(); controls_panel.setFixedWidth(300)
        controls_layout = QVBoxLayout(controls_panel); self.grid_widget = GridWidget()
        main_layout.addWidget(controls_panel); main_layout.addWidget(self.grid_widget)
        form_layout = QFormLayout()
        self.start_x = QSpinBox(); self.start_x.setRange(1, 1000); self.start_x.setValue(5)
        self.start_z = QSpinBox(); self.start_z.setRange(1, 1000); self.start_z.setValue(5)
        self.end_x = QSpinBox(); self.end_x.setRange(1, 1000); self.end_x.setValue(35)
        self.end_z = QSpinBox(); self.end_z.setRange(1, 1000); self.end_z.setValue(35)
        form_layout.addRow("Start X/Z:", self._create_pos_layout(self.start_x, self.start_z))
        form_layout.addRow("End X/Z:", self._create_pos_layout(self.end_x, self.end_z))
        self.spacing_input = QSpinBox(); self.spacing_input.setRange(0, 10); self.spacing_input.setValue(1)
        self.branching_input = QComboBox(); self.branching_input.addItems(["On", "Off"])
        self.fill_slider = QSlider(Qt.Horizontal); self.fill_slider.setRange(1, 100); self.fill_slider.setValue(20)
        self.fill_label = QLabel("20 %"); self.fill_slider.valueChanged.connect(lambda v: self.fill_label.setText(f"{v} %"))
        fill_layout = QHBoxLayout(); fill_layout.addWidget(self.fill_slider); fill_layout.addWidget(self.fill_label)
        self.seed_input = QSpinBox(); self.seed_input.setRange(-1, 2147483647)
        randomize_seed_button = QPushButton("🎲"); randomize_seed_button.clicked.connect(self.randomize_seed)
        seed_layout = QHBoxLayout(); seed_layout.addWidget(self.seed_input); seed_layout.addWidget(randomize_seed_button)
        self.randomize_seed()
        form_layout.addRow("Path Spacing:", self.spacing_input)
        form_layout.addRow("Branching (Dead Ends):", self.branching_input)
        form_layout.addRow("Path Density (%):", fill_layout)
        form_layout.addRow("Seed:", seed_layout)
        generate_button = QPushButton("Generate Path"); generate_button.clicked.connect(self.run_generation)
        controls_layout.addLayout(form_layout); controls_layout.addStretch(); controls_layout.addWidget(generate_button)
    def _create_pos_layout(self, x_spin, z_spin):
        layout = QHBoxLayout(); layout.addWidget(x_spin); layout.addWidget(z_spin); return layout
    def randomize_seed(self): self.seed_input.setValue(int(time.time()) + random.randint(0, 10000))
    def run_generation(self):
        start_pos = (self.start_x.value(), self.start_z.value()); end_pos = (self.end_x.value(), self.end_z.value())
        spacing = self.spacing_input.value(); branching = self.branching_input.currentText() == "On"
        fill_percent = self.fill_slider.value(); seed = self.seed_input.value()
        generator = PathGenerator(start_pos, end_pos, spacing, branching, fill_percent, seed)
        grid = generator.generate(); self.grid_widget.set_grid_data(grid)
if __name__ == "__main__":
    app = QApplication(sys.argv); window = MainWindow(); window.show(); window.run_generation(); sys.exit(app.exec_())
