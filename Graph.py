import sys
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QPushButton, QLabel, QFileDialog, QWidget, QComboBox, QHBoxLayout, QLineEdit, QMessageBox
)
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import matplotlib.pyplot as plt

class GraphApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Graph Creator")

        # Main widget and layout
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout()
        self.central_widget.setLayout(self.layout)

        # Widgets
        self.upload_button = QPushButton("Upload Excel File")
        self.upload_button.clicked.connect(self.upload_file)
        self.layout.addWidget(self.upload_button)

        self.file_label = QLabel("No file selected")
        self.layout.addWidget(self.file_label)

        self.column_selector_layout = QHBoxLayout()

        self.y_column_label = QLabel("Y-axis:")
        self.column_selector_layout.addWidget(self.y_column_label)
        self.y_column_combo = QComboBox()
        self.column_selector_layout.addWidget(self.y_column_combo)

        self.range_label = QLabel("Range (start:end):")
        self.column_selector_layout.addWidget(self.range_label)
        self.range_input = QLineEdit()
        self.column_selector_layout.addWidget(self.range_input)

        self.layout.addLayout(self.column_selector_layout)

        self.plot_button = QPushButton("Plot Graph")
        self.plot_button.clicked.connect(self.plot_graph)
        self.layout.addWidget(self.plot_button)

        self.add_signal_button = QPushButton("Add Signal")
        self.add_signal_button.clicked.connect(self.add_signal)
        self.layout.addWidget(self.add_signal_button)

        self.edit_button = QPushButton("Edit Range")
        self.edit_button.clicked.connect(self.edit_range)
        self.layout.addWidget(self.edit_button)

        self.save_button = QPushButton("Save Plot as Image")
        self.save_button.clicked.connect(self.save_image)
        self.layout.addWidget(self.save_button)

        self.canvas = FigureCanvas(Figure(figsize=(5, 3)))
        self.layout.addWidget(self.canvas)

        self.ax = self.canvas.figure.add_subplot(111)

        # Data storage
        self.data = None
        self.x_column = 'a'
        self.colors = plt.cm.tab10.colors  # Color cycle for differentiation
        self.plot_count = 0  # Track the number of signals added
        self.plot_signals = set()  # Set to track added signals

    def upload_file(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Open Excel File", "", "Excel Files (*.xlsx *.xls)", options=options
        )
        if file_path:
            self.file_label.setText(file_path)
            try:
                self.data = pd.read_excel(file_path)
                if self.x_column not in self.data.columns:
                    QMessageBox.critical(self, "Error", f"Column '{self.x_column}' not found in the Excel file.")
                    self.data = None
                    return
                self.update_column_selector()
            except Exception as e:
                self.file_label.setText(f"Error reading file: {str(e)}")

    def update_column_selector(self):
        if self.data is not None:
            self.y_column_combo.clear()
            columns = [col for col in self.data.columns if col != self.x_column]
            self.y_column_combo.addItems(columns)
            self.range_input.setText(f"0:{len(self.data) - 1}")

    def edit_range(self):
        """Updates the range for the plot when the user clicks the Edit Range button."""
        if self.data is not None:
            range_text = self.range_input.text()
            try:
                start, end = map(int, range_text.split(':'))
                if start < 0 or end >= len(self.data) or start > end:
                    raise ValueError
                QMessageBox.information(self, "Success", f"Range updated to: {start} to {end}")
                self.plot_graph(start, end)
            except Exception:
                QMessageBox.critical(self, "Error", "Invalid range. Please enter as 'start:end' where start and end are valid indices.")

    def add_signal(self):
        """Adds an additional signal to the plot if not already added."""
        if self.data is not None:
            y_column = self.y_column_combo.currentText()
            range_text = self.range_input.text()
            try:
                start, end = map(int, range_text.split(':'))
                if start < 0 or end >= len(self.data) or start > end:
                    raise ValueError

                # Check if the signal has already been added
                if y_column in self.plot_signals:
                    QMessageBox.warning(self, "Warning", f"Signal '{y_column}' has already been added.")
                    return

                # Plot the new signal with a different color
                self.ax.plot(
                    self.data[self.x_column].iloc[start:end + 1],
                    self.data[y_column].iloc[start:end + 1],
                    marker='o',
                    linewidth=0.1,  # Adjust line width
                    label=f"{y_column} (Signal {self.plot_count + 1})",
                    color=self.colors[self.plot_count % len(self.colors)]
                )

                self.plot_signals.add(y_column)  # Mark this signal as added
                self.plot_count += 1

                # Update the legend and grid
                self.ax.legend()
                self.ax.grid(True)


                self.canvas.draw()

            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to add signal: {str(e)}")

    def plot_graph(self, start=None, end=None):
        """Plots the graph with the selected range."""
        if self.data is not None:
            if start is None or end is None:
                range_text = self.range_input.text()
                try:
                    start, end = map(int, range_text.split(':'))
                    if start < 0 or end >= len(self.data) or start > end:
                        raise ValueError
                except Exception as e:
                    QMessageBox.critical(self, "Error", f"Failed to plot graph: {str(e)}")
                    return

            self.ax.clear()  # Clear the axes before redrawing
            self.plot_signals.clear()  # Reset the signals that were plotted
            self.plot_count = 0  # Reset plot count

            # Plot the first signal
            y_column = self.y_column_combo.currentText()
            self.ax.plot(
                self.data[self.x_column].iloc[start:end + 1],
                self.data[y_column].iloc[start:end + 1],
                marker='o', label=f"{y_column} (Signal 1)", color=self.colors[self.plot_count % len(self.colors)]
            )
            self.plot_signals.add(y_column)  # Mark the first signal as added
            self.plot_count += 1

            # Update title, legend, and grid
            self.ax.set_title(f"Graph of Signals vs {self.x_column}")
            self.ax.set_xlabel(self.x_column)
            self.ax.set_ylabel("Signals")
            self.ax.legend()
            self.ax.grid(True)

            self.canvas.draw()


    def save_image(self):
        file_path, _ = QFileDialog.getSaveFileName(self, "Save Plot as Image", "", "PNG Files (*.png);;JPEG Files (*.jpg);;All Files (*)")
        if file_path:
            try:
                self.canvas.figure.savefig(file_path)
                QMessageBox.information(self, "Success", f"Graph saved as image: {file_path}")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to save image: {str(e)}")


# Run the application
if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = GraphApp()
    main_window.show()
    sys.exit(app.exec_())
